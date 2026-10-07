/**
 * Coverage for woff2-decoder.js's WOFF2 transformed-glyf reconstruction
 * (reconstructGlyf) — the piece that turns WOFF2's compact "triplet"-encoded
 * glyph outlines back into standard sfnt `glyf`/`loca` bytes. The fixture
 * below is a tiny, deliberately minimal font (4 glyphs: an empty glyph, a
 * 3-point simple glyph with one off-curve point, and a 2-component composite
 * glyph) built with fontTools (the reference WOFF2 implementation this
 * decoder was ported from — see woff2-decoder.js's own top comment) and
 * saved with the glyf transform applied, so this test exercises the exact
 * real-world binary format rather than a hand-rolled approximation of it.
 * The decoded output was independently cross-checked glyph-for-glyph against
 * fontTools' own decode of both this fixture and several real downloaded
 * fonts (Roboto, Open Sans) during development; this test pins the same
 * fixture's expected values so a future regression is caught without needing
 * Python/fontTools available in this repo's own test run.
 */
import { describe, it, expect } from 'vitest';
import { isWoff2, decodeWoff2 } from '../src/fonts/woff2-decoder.js';
import { parseSfnt } from '../src/fonts/sfnt-parser.js';

// Glyph order: .notdef (simple, 4-point square), space (empty), triangle
// (simple, 3 points — the 3rd is off-curve), double (composite: two
// "triangle" components, one at (0,0) and one at (300,200)).
const FIXTURE_WOFF2_BASE64 = 'd09GMgABAAAAAAFMAAoAAAAAAtAAAAEGAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAABmAAPApMWQE2AiQDCgsKAAQgBWwHOhseAgCeBXaTG7U+rEZs2/XpYlMjIyiHbS/J16ppTVWLse8UVqAZ3xXuHXRfdHSb1tF9ekf0R8cQQTpMwYYECOOZOMP5Ojm+SU6BtUUUtTQAiizxFkWY25i4nFIQrWZ1UU6roKEQBRtwxYWN3Rs9GyAhYy4LzBUHaCpC6mK1Eu+vr+upAaCANhJa6GKIPghErSX3Bo1+DNanVvj7ly0JWYoTd+OAvJFuBYKX+Tuj/bcj/8DX9fkztF44GgR5oyKtUJU3BAAg0goBAkURmpIEAJJAcSlA6CFA0tKTRekIS9Pck6GDRbbMWpGOL1NfoA2+wFQJycECIpnLdVYU5EiR27c58cQwhgo=';

function readTableDirectory(buf) {
  const numTables = buf.readUInt16BE(4);
  const tables = new Map();
  for (let i = 0; i < numTables; i++) {
    const entryOffset = 12 + i * 16;
    const tag = buf.toString('latin1', entryOffset, entryOffset + 4);
    const offset = buf.readUInt32BE(entryOffset + 8);
    const length = buf.readUInt32BE(entryOffset + 12);
    tables.set(tag, buf.subarray(offset, offset + length));
  }
  return tables;
}

describe('WOFF2 transformed-glyf reconstruction', () => {
  const fixture = Buffer.from(FIXTURE_WOFF2_BASE64, 'base64');

  it('isWoff2 recognizes the signature', () => {
    expect(isWoff2(fixture)).toBe(true);
  });

  it('decodes without throwing and produces a valid sfnt', () => {
    const sfnt = decodeWoff2(fixture);
    const parsed = parseSfnt(sfnt);
    expect(parsed.numGlyphs).toBe(4);
    expect(parsed.glyphIdForCodepoint(32)).toBe(1); // space
    expect(parsed.glyphIdForCodepoint(0x25b2)).toBe(2); // triangle
    expect(parsed.advanceWidthForGlyph(0)).toBe(600);
  });

  it('reconstructs the simple .notdef glyph (4-point square) correctly', () => {
    const tables = readTableDirectory(decodeWoff2(fixture));
    const glyf = tables.get('glyf');
    const loca = tables.get('loca');
    // indexFormat 0 (short) — offsets are value/2, uint16 each; 5 entries for 4 glyphs.
    const offsets = [0, 1, 2, 3, 4].map((i) => loca.readUInt16BE(i * 2) * 2);
    const notdef = glyf.subarray(offsets[0], offsets[1]);

    expect(notdef.readInt16BE(0)).toBe(1); // numberOfContours
    expect([notdef.readInt16BE(2), notdef.readInt16BE(4), notdef.readInt16BE(6), notdef.readInt16BE(8)])
      .toEqual([50, 0, 450, 600]); // xMin, yMin, xMax, yMax
  });

  it('reconstructs the empty "space" glyph as zero-length', () => {
    const tables = readTableDirectory(decodeWoff2(fixture));
    const loca = tables.get('loca');
    const offsets = [0, 1, 2, 3, 4].map((i) => loca.readUInt16BE(i * 2) * 2);
    expect(offsets[2] - offsets[1]).toBe(0); // space (glyph 1) has no data
  });

  it('reconstructs the simple "triangle" glyph with correct on/off-curve flags and coordinates', () => {
    const tables = readTableDirectory(decodeWoff2(fixture));
    const glyf = tables.get('glyf');
    const loca = tables.get('loca');
    const offsets = [0, 1, 2, 3, 4].map((i) => loca.readUInt16BE(i * 2) * 2);
    const triangle = glyf.subarray(offsets[2], offsets[3]);

    let pos = 0;
    const numberOfContours = triangle.readInt16BE(pos); pos += 2;
    expect(numberOfContours).toBe(1);
    pos += 8; // bbox
    const endPtsOfContours = [triangle.readUInt16BE(pos)]; pos += 2;
    expect(endPtsOfContours).toEqual([2]); // 3 points (indices 0..2)
    const instructionLength = triangle.readUInt16BE(pos); pos += 2;
    expect(instructionLength).toBe(0);
    pos += instructionLength;

    const flags = [triangle.readUInt8(pos), triangle.readUInt8(pos + 1), triangle.readUInt8(pos + 2)];
    pos += 3;
    // on-curve bit (0x01): points 0 and 1 on-curve, point 2 off-curve.
    expect(flags.map((f) => f & 0x01)).toEqual([1, 1, 0]);

    const dx = [triangle.readInt16BE(pos), triangle.readInt16BE(pos + 2), triangle.readInt16BE(pos + 4)];
    pos += 6;
    const dy = [triangle.readInt16BE(pos), triangle.readInt16BE(pos + 2), triangle.readInt16BE(pos + 4)];

    // Absolute coordinates were (0,0), (500,0), (250,500) — reconstruct from
    // deltas the same way a real glyf consumer would, to confirm the x-block/
    // y-block split and delta-from-previous-point encoding are both correct.
    const xs = []; let x = 0; for (const d of dx) { x += d; xs.push(x); }
    const ys = []; let y = 0; for (const d of dy) { y += d; ys.push(y); }
    expect(xs).toEqual([0, 500, 250]);
    expect(ys).toEqual([0, 0, 500]);
  });

  it('reconstructs the composite "double" glyph with two correctly-offset components', () => {
    const tables = readTableDirectory(decodeWoff2(fixture));
    const glyf = tables.get('glyf');
    const loca = tables.get('loca');
    const offsets = [0, 1, 2, 3, 4].map((i) => loca.readUInt16BE(i * 2) * 2);
    const double = glyf.subarray(offsets[3], offsets[4]);

    const numberOfContours = double.readInt16BE(0);
    expect(numberOfContours).toBe(-1); // composite marker

    // Two components, each: flags(u16) + glyphIndex(u16) + x,y as bytes
    // (ARG_1_AND_2_ARE_WORDS not set here, since offsets fit in a signed byte).
    let pos = 10; // past numberOfContours + bbox
    const components = [];
    for (let i = 0; i < 2; i++) {
      const flags = double.readUInt16BE(pos); pos += 2;
      const glyphIndex = double.readUInt16BE(pos); pos += 2;
      const argWords = (flags & 0x0001) !== 0;
      let x;
      let y;
      if (argWords) { x = double.readInt16BE(pos); y = double.readInt16BE(pos + 2); pos += 4; }
      else { x = double.readInt8(pos); y = double.readInt8(pos + 1); pos += 2; }
      components.push({ glyphIndex, x, y });
      const moreComponents = (flags & 0x0020) !== 0;
      if (!moreComponents) break;
    }

    expect(components).toEqual([
      { glyphIndex: 2, x: 0, y: 0 },
      { glyphIndex: 2, x: 300, y: 200 },
    ]);
  });
});
