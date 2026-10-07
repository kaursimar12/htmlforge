import { describe, expect, it } from 'vitest';
import { deflateSync } from 'node:zlib';
import { decodeWoff, isWoff } from '../src/fonts/woff-decoder.js';
import { parseSfnt } from '../src/fonts/sfnt-parser.js';
import { buildTestFont } from './helpers/sfnt.js';

function u16be(v) { const b = Buffer.alloc(2); b.writeUInt16BE(v, 0); return b; }
function u32be(v) { const b = Buffer.alloc(4); b.writeUInt32BE(v >>> 0, 0); return b; }

// Repackages an already-built sfnt buffer as a WOFF1 file — reads the
// sfnt's own table directory (12-byte header + 16-byte records) to recover
// each table's bytes, then re-encodes them per the WOFF1 container format
// (each table independently zlib-deflated), so decodeWoff has a real,
// self-consistent WOFF1 file to reconstruct from.
function wrapAsWoff1(sfnt) {
  const numTables = sfnt.readUInt16BE(4);
  const tables = [];
  for (let i = 0; i < numTables; i++) {
    const recordOffset = 12 + i * 16;
    const tag = sfnt.toString('latin1', recordOffset, recordOffset + 4);
    const offset = sfnt.readUInt32BE(recordOffset + 8);
    const length = sfnt.readUInt32BE(recordOffset + 12);
    tables.push({ tag, data: sfnt.subarray(offset, offset + length) });
  }

  const entries = tables.map(({ tag, data }) => {
    const compressed = deflateSync(data);
    // Only use the compressed form when it's actually smaller — WOFF1
    // signals "stored raw" by compLength === origLength (see
    // woff-decoder.js's own decodeWoff), so a tiny table where zlib
    // overhead outgrows the raw bytes must be stored raw to stay decodable.
    const useCompressed = compressed.length < data.length;
    return { tag, compData: useCompressed ? compressed : data, origLength: data.length };
  });

  const headerSize = 44;
  const dirSize = entries.length * 20;
  let cursor = headerSize + dirSize;
  const laidOut = entries.map((e) => {
    const offset = cursor;
    cursor += e.compData.length;
    return { ...e, offset };
  });

  const header = Buffer.concat([
    u32be(0x774f4646), // 'wOFF'
    u32be(0x00010000), // flavor
    u32be(cursor),      // length
    u16be(laidOut.length),
    u16be(0),           // reserved
    u32be(0),           // totalSfntSize (unused by decodeWoff)
    u16be(1), u16be(0), // version
    u32be(0), u32be(0), u32be(0), // meta
    u32be(0), u32be(0),           // priv
  ]);

  const dir = Buffer.concat(laidOut.map((e) => Buffer.concat([
    Buffer.from(e.tag, 'latin1'),
    u32be(e.offset),
    u32be(e.compData.length),
    u32be(e.origLength),
    u32be(0), // origChecksum — decodeWoff recomputes rather than trusting this
  ])));

  const body = Buffer.concat(laidOut.map((e) => e.compData));
  return Buffer.concat([header, dir, body]);
}

describe('decodeWoff', () => {
  it('recognizes a WOFF1 buffer via isWoff', () => {
    const woff = wrapAsWoff1(buildTestFont());
    expect(isWoff(woff)).toBe(true);
    expect(isWoff(Buffer.from('not a woff'))).toBe(false);
  });

  it('reconstructs an sfnt whose parsed metrics match the original font', () => {
    const original = buildTestFont({ unitsPerEm: 1000, ascent: 800, descent: -200, capHeight: 700 });
    const woff = wrapAsWoff1(original);
    const reconstructed = decodeWoff(woff);

    const originalParsed = parseSfnt(original);
    const reconstructedParsed = parseSfnt(reconstructed);
    expect(reconstructedParsed.unitsPerEm).toBe(originalParsed.unitsPerEm);
    expect(reconstructedParsed.ascent).toBe(originalParsed.ascent);
    expect(reconstructedParsed.descent).toBe(originalParsed.descent);
    expect(reconstructedParsed.capHeight).toBe(originalParsed.capHeight);
    expect(reconstructedParsed.glyphIdForCodepoint(0x41)).toBe(originalParsed.glyphIdForCodepoint(0x41));
    expect(reconstructedParsed.advanceWidthForGlyph(1)).toBe(originalParsed.advanceWidthForGlyph(1));
  });

  it('throws on a buffer with the wrong signature', () => {
    expect(() => decodeWoff(Buffer.alloc(50))).toThrow(/signature/);
  });
});
