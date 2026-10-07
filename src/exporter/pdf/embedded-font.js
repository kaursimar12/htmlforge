/**
 * Embedded (non-Standard-14) font → PDF composite-font object description —
 * the CID/embedded-font sibling of images.js's buildImageXObject. Builds a
 * Type0/CIDFontType2/FontDescriptor/FontFile2(/ToUnicode) resource
 * description with object-number PLACEHOLDERS (`true`), exactly like
 * images.js's `SMask: true` placeholder for a PNG's alpha sibling — real
 * object numbers don't exist yet at build time (pdf-writer.js is the only
 * place that ever allocates them), so every reference from a parent dict to
 * a not-yet-numbered child object is left as a placeholder here and resolved
 * into a real `ref(objNum)` by pdf-writer.js's own embedded-font write pass.
 *
 * Deliberately NOT subsetted — the whole original font program is embedded
 * via FontFile2 as-is. Subsetting only reduces file size; it doesn't change
 * what glyphs render, so it's a size optimization for later, not a
 * prerequisite for correct output (see the project discussion this
 * implements). `CIDToGIDMap: Identity` is exactly why unsubsetted embedding
 * needs no glyph remapping at all: the CIDs shown in the content stream
 * (Identity-H encoded 2-byte codes) are chosen to BE the font's own real
 * glyph IDs, so "CID equals GID" is true by construction, not by luck.
 */
import { deflateSync } from 'node:zlib';
import { name } from './pdf-writer.js';

function scaleToThousandUpm(value, unitsPerEm) {
  return Math.round((value * 1000) / unitsPerEm);
}

// A CID's advance width only needs listing in `/W` for glyphs this document
// can actually reach via the font's own cmap — a font's total glyph count
// (ligatures, unmapped alternates, etc.) is usually far larger than its
// Unicode-reachable subset, so this stays proportional to real coverage
// rather than the font's raw glyph count.
function buildWidthArray(sfnt) {
  const scale = 1000 / sfnt.unitsPerEm;
  const pairs = [...new Set(sfnt.cmap.values())]
    .map((gid) => [gid, Math.round(sfnt.advanceWidthForGlyph(gid) * scale)])
    .sort((a, b) => a[0] - b[0]);

  const w = [];
  let i = 0;
  while (i < pairs.length) {
    const startGid = pairs[i][0];
    const widths = [pairs[i][1]];
    let j = i + 1;
    while (j < pairs.length && pairs[j][0] === pairs[j - 1][0] + 1) {
      widths.push(pairs[j][1]);
      j++;
    }
    w.push(startGid, widths);
    i = j;
  }
  return w;
}

// A minimal bfrange-based ToUnicode CMap — one entry per glyph actually
// reachable via cmap, mapping that glyph's CID (== its own glyph ID, per the
// Identity CIDToGIDMap above) back to the Unicode codepoint that produced
// it. Without this, copy-pasting or text-searching the PDF's own embedded-
// font runs would resolve to nothing (a Type0/Identity-H font has no
// encoding a PDF reader could otherwise reverse), even though the glyphs
// render correctly either way.
function buildToUnicodeCMap(sfnt) {
  const entries = [];
  for (const [codepoint, glyphId] of sfnt.cmap.entries()) {
    const hex4 = (n) => n.toString(16).padStart(4, '0');
    // Codepoints beyond the BMP need a UTF-16 surrogate pair in the CMap's
    // own destination string, not a raw 4-hex-digit value.
    const dst = codepoint > 0xffff
      ? hex4(0xd800 + ((codepoint - 0x10000) >> 10)) + hex4(0xdc00 + ((codepoint - 0x10000) & 0x3ff))
      : hex4(codepoint);
    entries.push(`<${hex4(glyphId)}> <${dst}>`);
  }
  const BATCH = 100; // bfchar blocks are capped at 100 entries per the CMap spec
  const blocks = [];
  for (let i = 0; i < entries.length; i += BATCH) {
    blocks.push(`${entries.length - i < BATCH ? entries.length - i : BATCH} beginbfchar\n${entries.slice(i, i + BATCH).join('\n')}\nendbfchar`);
  }
  const body = `/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /Adobe-Identity-UCS def
/CMapType 2 def
1 begincodespacerange
<0000> <FFFF>
endcodespacerange
${blocks.join('\n')}`;
  return Buffer.from(`/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n${body}\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`, 'latin1');
}

/**
 * Builds the resource description for one embedded font, given a resolved
 * registry entry (`{sfnt, buffer, family, weight, italic}` — see
 * fonts/resolver.js) and the resourceName it should be drawn under (e.g.
 * "EF1" — a distinct namespace from the standard-14 fonts' "F1"/"F2", since
 * these are never the same kind of PDF object). Returns undefined (+ warns)
 * rather than throwing if the font can't be embedded — same recoverable
 * contract as buildImageXObject.
 */
export function buildEmbeddedFontDescriptor(entry, resourceName, warn) {
  const { sfnt, buffer, family, weight, italic } = entry;
  try {
    const scale = 1000 / sfnt.unitsPerEm;
    const postScriptName = `${family.replace(/\s+/g, '')}-${weight}${italic ? 'Italic' : ''}`;
    const compressed = deflateSync(buffer);

    let flags = 0x20; // Nonsymbolic — this is ordinary Latin text, not a symbol font
    if (italic) flags |= 0x40;
    if (weight >= 700) flags |= 0x40000; // ForceBold

    const fontFile = {
      data: compressed,
      dict: { Filter: name('FlateDecode'), Length1: buffer.length },
    };
    const fontDescriptor = {
      dict: {
        Type: name('FontDescriptor'),
        FontName: name(postScriptName),
        Flags: flags,
        FontBBox: sfnt.bboxFontUnits.map((v) => scaleToThousandUpm(v, sfnt.unitsPerEm)),
        ItalicAngle: sfnt.italicAngleDeg,
        Ascent: scaleToThousandUpm(sfnt.ascent, sfnt.unitsPerEm),
        Descent: scaleToThousandUpm(sfnt.descent, sfnt.unitsPerEm),
        CapHeight: scaleToThousandUpm(sfnt.capHeight, sfnt.unitsPerEm),
        StemV: weight >= 700 ? 120 : 80, // no direct measurement available; common embedder default
        FontFile2: true, // placeholder — pdf-writer.js swaps for a real ref()
      },
      fontFile,
    };
    const descendantFont = {
      dict: {
        Type: name('Font'),
        Subtype: name('CIDFontType2'),
        BaseFont: name(postScriptName),
        CIDSystemInfo: { Registry: 'Adobe', Ordering: 'Identity', Supplement: 0 },
        FontDescriptor: true, // placeholder
        DW: 1000,
        W: buildWidthArray(sfnt),
        CIDToGIDMap: name('Identity'),
      },
      fontDescriptor,
    };
    const toUnicodeData = buildToUnicodeCMap(sfnt);

    return {
      resourceName,
      dict: {
        Type: name('Font'),
        Subtype: name('Type0'),
        BaseFont: name(postScriptName),
        Encoding: name('Identity-H'),
        DescendantFonts: true, // placeholder — becomes [ref(objNum)]
        ToUnicode: true, // placeholder
      },
      descendantFont,
      toUnicode: { data: toUnicodeData, dict: {} },
      // Carried through for content-stream drawing (page-flow.js's
      // drawEmbeddedText) — glyph lookup and unitsPerEm-based sizing don't
      // need any PDF-object concept, just the parsed font data itself.
      glyphIdForCodepoint: (cp) => sfnt.glyphIdForCodepoint(cp),
      unitsPerEm: sfnt.unitsPerEm,
    };
  } catch (err) {
    warn?.('EMBEDDED_FONT_FAILED', `Cannot embed font "${family}": ${err instanceof Error ? err.message : String(err)}`);
    return undefined;
  }
}
