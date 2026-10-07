/**
 * sfnt (TrueType/OpenType) table parser — reads just enough of a font
 * program to embed it in a PDF CIDFontType2 unsubsetted (see
 * exporter/pdf/embedded-font.js): the table directory, `cmap` (Unicode →
 * glyph ID), `hmtx`/`hhea` (per-glyph advance widths, for the CIDFont's own
 * `/W` array), and `head`/`OS2`/`post` (FontDescriptor metrics — ascent/
 * descent/capHeight/italicAngle/bbox/bold+italic flags).
 *
 * Deliberately does NOT touch `glyf`/`loca`/CFF — embedding the whole
 * original font program via `FontFile2` needs none of that; a real font's
 * own outlines are opaque bytes to this parser, copied through as-is by the
 * caller. This also means the same parser works unmodified for both
 * TrueType-flavored (`glyf`) and CFF-flavored (`OTTO`) OpenType fonts, since
 * `cmap`/`hmtx`/`head`/`OS2`/`post` are identical table formats in both.
 */
import { ByteReader } from '../utils/bytes.js';

const SFNT_TRUETYPE = 0x00010000;
const SFNT_OTTO = 0x4f54544f; // 'OTTO'
const SFNT_TRUE = 0x74727565; // 'true' (older Mac TrueType)

export function isSfnt(buffer) {
  if (buffer.length < 4) return false;
  const tag = buffer.readUInt32BE(0);
  return tag === SFNT_TRUETYPE || tag === SFNT_OTTO || tag === SFNT_TRUE;
}

function readTableDirectory(r) {
  const tag = r.u32be();
  if (tag !== SFNT_TRUETYPE && tag !== SFNT_OTTO && tag !== SFNT_TRUE) {
    throw new Error('Not a TrueType/OpenType font (unrecognized sfnt version tag)');
  }
  const numTables = r.u16be();
  r.skip(6); // searchRange, entrySelector, rangeShift
  const tables = new Map();
  for (let i = 0; i < numTables; i++) {
    const name = r.string(4, 'latin1');
    r.skip(4); // checksum
    const offset = r.u32be();
    const length = r.u32be();
    tables.set(name, { offset, length });
  }
  return tables;
}

function parseHead(r, table) {
  r.seek(table.offset + 18);
  const unitsPerEm = r.u16be();
  r.seek(table.offset + 36);
  const xMin = r.i16be();
  const yMin = r.i16be();
  const xMax = r.i16be();
  const yMax = r.i16be();
  const macStyle = r.u16be();
  return { unitsPerEm, bbox: [xMin, yMin, xMax, yMax], bold: (macStyle & 0x1) !== 0, italic: (macStyle & 0x2) !== 0 };
}

function parseHhea(r, table) {
  r.seek(table.offset + 4);
  const ascent = r.i16be();
  const descent = r.i16be();
  r.skip(2 + 2); // lineGap, advanceWidthMax
  r.seek(table.offset + 34);
  const numberOfHMetrics = r.u16be();
  return { ascent, descent, numberOfHMetrics };
}

function parseMaxp(r, table) {
  r.seek(table.offset + 4);
  return { numGlyphs: r.u16be() };
}

// `hmtx`: `numberOfHMetrics` (advanceWidth uint16, lsb int16) pairs, then any
// remaining glyph beyond that count shares the last advanceWidth (only its
// own lsb is stored) — the standard sfnt space-saving convention for fonts
// where most glyphs share one advance width (monospace, or a large tail of
// rarely-distinct-width glyphs).
function parseHmtx(r, table, numberOfHMetrics, numGlyphs) {
  const widths = new Uint16Array(numGlyphs);
  r.seek(table.offset);
  let last = 0;
  for (let i = 0; i < numberOfHMetrics; i++) {
    last = r.u16be();
    r.skip(2); // lsb
    widths[i] = last;
  }
  for (let i = numberOfHMetrics; i < numGlyphs; i++) widths[i] = last;
  return widths;
}

function parseOS2(r, table) {
  if (!table) return {};
  r.seek(table.offset + 4);
  const usWeightClass = r.u16be();
  r.seek(table.offset + 62);
  const fsSelection = r.u16be();
  let capHeight;
  // sCapHeight only exists in OS/2 version >= 2 (table itself is long enough
  // to hold it at offset 88); a v0/v1 table simply doesn't have this field.
  if (table.length >= 90) {
    r.seek(table.offset + 88);
    capHeight = r.i16be();
  }
  return {
    weightClass: usWeightClass,
    italic: (fsSelection & 0x1) !== 0,
    bold: (fsSelection & 0x20) !== 0,
    capHeight,
  };
}

function parsePost(r, table) {
  if (!table) return {};
  r.seek(table.offset + 4);
  // italicAngle is a 16.16 fixed-point signed value.
  const raw = r.u32be();
  const signed = raw >= 0x80000000 ? raw - 0x100000000 : raw;
  return { italicAngleDeg: signed / 65536 };
}

// ─── cmap ───────────────────────────────────────────────────────────────

// Picks the best available Unicode subtable: prefer full-repertoire format
// 12 (Windows/Unicode, platform 3 encoding 10) over BMP-only format 4
// (platform 3 encoding 1), falling back to any Unicode-platform (0) entry.
function selectCmapSubtable(r, table) {
  r.seek(table.offset);
  r.skip(2); // version
  const numTables = r.u16be();
  const candidates = [];
  for (let i = 0; i < numTables; i++) {
    const platformID = r.u16be();
    const encodingID = r.u16be();
    const offset = r.u32be();
    candidates.push({ platformID, encodingID, offset: table.offset + offset });
  }
  const rank = (c) => {
    if (c.platformID === 3 && c.encodingID === 10) return 0;
    if (c.platformID === 3 && c.encodingID === 1) return 1;
    if (c.platformID === 0) return 2;
    return 3;
  };
  candidates.sort((a, b) => rank(a) - rank(b));
  return candidates[0];
}

function parseCmapFormat4(r, start, cmap) {
  r.seek(start);
  const format = r.u16be();
  if (format !== 4) return;
  r.skip(2 + 2); // length, language
  const segCountX2 = r.u16be();
  const segCount = segCountX2 / 2;
  r.skip(6); // searchRange, entrySelector, rangeShift
  const endCodes = Array.from({ length: segCount }, () => r.u16be());
  r.skip(2); // reservedPad
  const startCodes = Array.from({ length: segCount }, () => r.u16be());
  const idDeltas = Array.from({ length: segCount }, () => r.i16be());
  const idRangeOffsetPos = r.offset;
  const idRangeOffsets = Array.from({ length: segCount }, () => r.u16be());

  for (let i = 0; i < segCount; i++) {
    const endCode = endCodes[i];
    const startCode = startCodes[i];
    if (startCode === 0xffff && endCode === 0xffff) continue; // terminator segment
    for (let c = startCode; c <= endCode; c++) {
      let glyphId;
      if (idRangeOffsets[i] === 0) {
        glyphId = (c + idDeltas[i]) & 0xffff;
      } else {
        const glyphIndexAddress = idRangeOffsetPos + i * 2 + idRangeOffsets[i] + (c - startCode) * 2;
        r.seek(glyphIndexAddress);
        const raw = r.u16be();
        glyphId = raw === 0 ? 0 : (raw + idDeltas[i]) & 0xffff;
      }
      if (glyphId !== 0 && !cmap.has(c)) cmap.set(c, glyphId);
    }
  }
}

function parseCmapFormat12(r, start, cmap) {
  r.seek(start);
  const format = r.u16be();
  if (format !== 12) return;
  r.skip(2 + 4 + 4); // reserved, length, language
  const numGroups = r.u32be();
  for (let i = 0; i < numGroups; i++) {
    const startCharCode = r.u32be();
    const endCharCode = r.u32be();
    const startGlyphID = r.u32be();
    for (let c = startCharCode; c <= endCharCode; c++) {
      const glyphId = startGlyphID + (c - startCharCode);
      if (!cmap.has(c)) cmap.set(c, glyphId);
    }
  }
}

function parseCmap(r, table) {
  const cmap = new Map();
  const subtable = selectCmapSubtable(r, table);
  if (!subtable) return cmap;
  r.seek(subtable.offset);
  const format = r.u16be();
  if (format === 4) parseCmapFormat4(r, subtable.offset, cmap);
  else if (format === 12) parseCmapFormat12(r, subtable.offset, cmap);
  // Other formats (0, 2, 6, 13, 14 — legacy/CJK-specific/variation-selector)
  // are rare for the Latin webfonts this feature targets; left unsupported
  // rather than guessed at, same as this project's other documented format bounds.
  return cmap;
}

/**
 * Parses a raw sfnt buffer (TrueType or CFF-flavored OpenType — already
 * unwrapped from WOFF/WOFF2 if the source was compressed) into the data an
 * embedded-CIDFontType2 PDF object needs. Throws only on a truly unreadable
 * buffer (wrong magic, truncated past a required table) — callers (the font
 * resolver) treat that as "can't embed this font" and warn, never crash the
 * export, matching this project's existing image-asset error handling.
 */
export function parseSfnt(buffer) {
  const r = new ByteReader(buffer);
  const tables = readTableDirectory(r);

  const required = ['head', 'hhea', 'hmtx', 'maxp', 'cmap'];
  for (const name of required) {
    if (!tables.has(name)) throw new Error(`Font is missing required table "${name}"`);
  }

  const head = parseHead(r, tables.get('head'));
  const hhea = parseHhea(r, tables.get('hhea'));
  const maxp = parseMaxp(r, tables.get('maxp'));
  const widths = parseHmtx(r, tables.get('hmtx'), hhea.numberOfHMetrics, maxp.numGlyphs);
  const os2 = parseOS2(r, tables.get('OS/2'));
  const post = parsePost(r, tables.get('post'));
  const cmap = parseCmap(r, tables.get('cmap'));

  return {
    unitsPerEm: head.unitsPerEm,
    bboxFontUnits: head.bbox,
    ascent: hhea.ascent,
    descent: hhea.descent,
    // OS/2's sCapHeight is the accurate value when present (version >= 2);
    // ~70% of ascent is a common approximation used by other font tooling
    // when a font's OS/2 table predates that field.
    capHeight: os2.capHeight ?? Math.round(hhea.ascent * 0.7),
    italicAngleDeg: post.italicAngleDeg ?? 0,
    bold: os2.bold ?? head.bold,
    italic: os2.italic ?? head.italic,
    weightClass: os2.weightClass,
    numGlyphs: maxp.numGlyphs,
    // An `fvar` table means this is a variable font — its glyf/loca/hmtx
    // only ever hold ONE default instance's outlines; a caller resolving
    // multiple weights from the same variable source (e.g. two @font-face
    // rules that happen to share one file) would otherwise silently embed
    // IDENTICAL, non-differentiated glyph data under both weight keys. See
    // fonts/stylesheet-resolver.js and google-fonts-fallback.js, which both
    // check this to avoid registering a fake "Bold"/"Italic" from a variable
    // source this project has no instantiation engine to actually resolve.
    isVariable: tables.has('fvar'),
    cmap,
    glyphIdForCodepoint(cp) { return cmap.get(cp); },
    advanceWidthForGlyph(gid) { return widths[gid] ?? widths[widths.length - 1] ?? 0; },
  };
}
