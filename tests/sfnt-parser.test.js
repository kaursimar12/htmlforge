import { describe, expect, it } from 'vitest';
import { parseSfnt, isSfnt } from '../src/fonts/sfnt-parser.js';
import { buildTestFont } from './helpers/sfnt.js';

describe('parseSfnt', () => {
  it('recognizes a valid sfnt buffer via isSfnt', () => {
    const font = buildTestFont();
    expect(isSfnt(font)).toBe(true);
    expect(isSfnt(Buffer.from('not a font'))).toBe(false);
  });

  it('parses unitsPerEm, ascent/descent/capHeight, and bbox from head/hhea/OS2', () => {
    const font = buildTestFont({ unitsPerEm: 2048, ascent: 1600, descent: -400, capHeight: 1400 });
    const parsed = parseSfnt(font);
    expect(parsed.unitsPerEm).toBe(2048);
    expect(parsed.ascent).toBe(1600);
    expect(parsed.descent).toBe(-400);
    expect(parsed.capHeight).toBe(1400);
    expect(parsed.bboxFontUnits).toEqual([0, 0, 2048, 2048]);
  });

  it('maps codepoints to glyph IDs via cmap format 4', () => {
    const font = buildTestFont({ glyphs: [{ cp: 0x41, width: 500 }, { cp: 0x42, width: 600 }] });
    const parsed = parseSfnt(font);
    expect(parsed.glyphIdForCodepoint(0x41)).toBe(1);
    expect(parsed.glyphIdForCodepoint(0x42)).toBe(2);
    expect(parsed.glyphIdForCodepoint(0x43)).toBeUndefined();
  });

  it('reads per-glyph advance widths from hmtx', () => {
    const font = buildTestFont({ glyphs: [{ cp: 0x41, width: 500 }, { cp: 0x42, width: 600 }] });
    const parsed = parseSfnt(font);
    expect(parsed.advanceWidthForGlyph(1)).toBe(500);
    expect(parsed.advanceWidthForGlyph(2)).toBe(600);
  });

  it('throws a descriptive error on a truncated/malformed buffer rather than silently misparsing', () => {
    const font = buildTestFont();
    expect(() => parseSfnt(font.subarray(0, 20))).toThrow();
  });

  it('flags a font with an fvar table as isVariable, and one without as not', () => {
    expect(parseSfnt(buildTestFont({ variable: true })).isVariable).toBe(true);
    expect(parseSfnt(buildTestFont({ variable: false })).isVariable).toBe(false);
  });
});
