/**
 * Builds a minimal, valid, synthetic sfnt buffer (head/hhea/maxp/hmtx/cmap/
 * OS2/post — everything sfnt-parser.js reads; no glyf/loca/CFF, since
 * neither this project's sfnt-parser nor its embedded-font PDF writer ever
 * touch glyph outlines) for exercising sfnt-parser.js, woff-decoder.js, and
 * woff2-decoder.js without depending on a real font file.
 */
import { buildSfnt } from '../../src/fonts/woff-decoder.js';

function u16be(v) { const b = Buffer.alloc(2); b.writeUInt16BE(v & 0xffff, 0); return b; }
function i16be(v) { const b = Buffer.alloc(2); b.writeInt16BE(v, 0); return b; }
function u32be(v) { const b = Buffer.alloc(4); b.writeUInt32BE(v >>> 0, 0); return b; }

export function buildTestFont({
  unitsPerEm = 1000,
  ascent = 800,
  descent = -200,
  capHeight = 700,
  glyphs = [{ cp: 0x41, width: 500 }, { cp: 0x42, width: 600 }],
  variable = false, // adds a (dummy, unparsed) `fvar` table — sfnt-parser.js only checks for its presence
} = {}) {
  const numGlyphs = glyphs.length + 1; // +1 for glyph 0 (.notdef)

  const head = Buffer.concat([
    u16be(1), u16be(0),          // version
    u32be(0x00010000),           // fontRevision
    u32be(0),                    // checkSumAdjustment (recomputed by buildSfnt)
    u32be(0x5f0f3cf5),           // magicNumber
    u16be(0),                    // flags
    u16be(unitsPerEm),
    Buffer.alloc(8), Buffer.alloc(8), // created, modified
    i16be(0), i16be(0), i16be(unitsPerEm), i16be(unitsPerEm), // xMin,yMin,xMax,yMax
    u16be(0),                    // macStyle
    u16be(8), i16be(2), i16be(0), // lowestRecPPEM, fontDirectionHint, indexToLocFormat
    i16be(0),                    // glyphDataFormat
  ]);

  const hhea = Buffer.concat([
    u16be(1), u16be(0),
    i16be(ascent), i16be(descent), i16be(0),
    u16be(0), i16be(0), i16be(0), i16be(0),
    i16be(0), i16be(0), i16be(0),
    Buffer.alloc(8),
    i16be(0),
    u16be(numGlyphs), // numberOfHMetrics — every glyph gets its own hmtx entry, simplest to construct
  ]);

  const maxp = Buffer.concat([u32be(0x00005000), u16be(numGlyphs)]);

  const hmtxParts = [u16be(0), i16be(0)]; // glyph 0 (.notdef)
  for (const g of glyphs) hmtxParts.push(u16be(g.width), i16be(0));
  const hmtx = Buffer.concat(hmtxParts);

  // cmap: one Windows/Unicode BMP (platform 3, encoding 1) format-4 subtable,
  // one segment per glyph (mapping cp -> glyphId 1,2,3...) plus the
  // mandatory 0xFFFF terminator segment.
  const segCount = glyphs.length + 1;
  const endCodes = glyphs.map((g) => g.cp).concat([0xffff]);
  const startCodes = glyphs.map((g) => g.cp).concat([0xffff]);
  const idDeltas = glyphs.map((g, i) => (i + 1) - g.cp).concat([1]);
  const idRangeOffsets = new Array(segCount).fill(0);

  let format4 = Buffer.concat([
    u16be(4), u16be(0), u16be(0), // format, length (patched below), language
    u16be(segCount * 2),
    u16be(0), u16be(0), u16be(0), // searchRange/entrySelector/rangeShift — unused by this project's reader
    ...endCodes.map(u16be),
    u16be(0), // reservedPad
    ...startCodes.map(u16be),
    ...idDeltas.map(i16be),
    ...idRangeOffsets.map(u16be),
  ]);
  format4.writeUInt16BE(format4.length, 2);

  const cmapHeader = Buffer.concat([u16be(0), u16be(1), u16be(3), u16be(1), u32be(12)]);
  const cmap = Buffer.concat([cmapHeader, format4]);

  // OS/2 version 4 — allocated as one zeroed block, only the specific
  // offsets sfnt-parser.js actually reads are patched in; every field this
  // project ignores is left zero rather than laid out field-by-field.
  const os2 = Buffer.alloc(96);
  os2.writeUInt16BE(4, 0); // version
  os2.writeUInt16BE(400, 4); // usWeightClass
  os2.writeUInt16BE(0x40, 62); // fsSelection — bit 6 (0x40) = REGULAR (neither bold nor italic)
  os2.writeInt16BE(capHeight, 88); // sCapHeight

  const post = Buffer.concat([u32be(0x00020000), u32be(0), Buffer.alloc(8)]); // version 2, italicAngle 0

  const tables = [
    { tag: 'head', data: head },
    { tag: 'hhea', data: hhea },
    { tag: 'maxp', data: maxp },
    { tag: 'hmtx', data: hmtx },
    { tag: 'cmap', data: cmap },
    { tag: 'OS/2', data: os2 },
    { tag: 'post', data: post },
  ];
  if (variable) tables.push({ tag: 'fvar', data: Buffer.alloc(16) });

  return buildSfnt(0x00010000, tables);
}
