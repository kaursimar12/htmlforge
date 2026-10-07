/**
 * WOFF (version 1) container decoder — reconstructs a plain sfnt buffer
 * (the same shape sfnt-parser.js/embedded-font.js already consume for a
 * plain .ttf/.otf) from a .woff file's own per-table zlib-compressed
 * layout. Unlike WOFF2 (see woff2-decoder.js), WOFF1 never transforms
 * `glyf`/`loca` — each table is independently zlib-`deflate`d (or stored
 * raw, if compression didn't shrink it), so reconstruction is just "inflate
 * each table, lay them out as a normal sfnt directory" — no glyph-outline
 * re-encoding algorithm needed, which is why this file is a fraction of the
 * size woff2-decoder.js's transform handling requires.
 */
import { inflateSync } from 'node:zlib';
import { ByteReader, ByteWriter } from '../utils/bytes.js';

const WOFF_SIGNATURE = 0x774f4646; // 'wOFF'

export function isWoff(buffer) {
  return buffer.length >= 4 && buffer.readUInt32BE(0) === WOFF_SIGNATURE;
}

function tableChecksum(buffer) {
  let sum = 0;
  for (let i = 0; i < buffer.length; i += 4) {
    const word = ((buffer[i] ?? 0) << 24) | ((buffer[i + 1] ?? 0) << 16) | ((buffer[i + 2] ?? 0) << 8) | (buffer[i + 3] ?? 0);
    sum = (sum + word) >>> 0;
  }
  return sum >>> 0;
}

/**
 * Decodes a WOFF1 buffer into a standalone sfnt buffer, ready to hand to
 * sfnt-parser.js/embedded-font.js exactly like a native .ttf/.otf would be.
 * Throws on a malformed/truncated file — callers (fonts/resolver.js) catch
 * this the same as any other unembeddable-font condition and warn rather
 * than fail the export.
 */
export function decodeWoff(buffer) {
  const r = new ByteReader(buffer);
  const signature = r.u32be();
  if (signature !== WOFF_SIGNATURE) throw new Error('Not a WOFF file (bad signature)');
  const flavor = r.u32be();
  r.skip(4); // length (total WOFF file size) — unused, we trust the actual buffer
  const numTables = r.u16be();
  r.skip(2); // reserved
  r.skip(4); // totalSfntSize — recomputed below rather than trusted
  r.skip(2 + 2); // majorVersion, minorVersion
  r.skip(4 + 4 + 4); // metaOffset, metaLength, metaOrigLength
  r.skip(4 + 4); // privOffset, privLength

  const entries = [];
  for (let i = 0; i < numTables; i++) {
    const tag = r.string(4, 'latin1');
    const offset = r.u32be();
    const compLength = r.u32be();
    const origLength = r.u32be();
    r.skip(4); // origChecksum — recomputed below rather than trusted
    entries.push({ tag, offset, compLength, origLength });
  }

  const tables = entries.map(({ tag, offset, compLength, origLength }) => {
    const compressed = buffer.subarray(offset, offset + compLength);
    // WOFF1 stores a table uncompressed (raw passthrough) whenever
    // compression wouldn't have shrunk it — signaled by compLength ===
    // origLength, per the WOFF1 spec, rather than any magic-byte sniff.
    const data = compLength === origLength ? compressed : inflateSync(compressed);
    if (data.length !== origLength) throw new Error(`Table "${tag}" decompressed to ${data.length} bytes, expected ${origLength}`);
    return { tag, data };
  });

  return buildSfnt(flavor, tables);
}

/**
 * Lays `tables` out as a standard sfnt file — table directory entries
 * sorted by tag (the conventional, safest order; some strict parsers expect
 * it even though the spec only requires the directory to correctly point
 * at each table, in any order), each table padded to a 4-byte boundary, and
 * `head`'s own `checksumAdjustment` field recomputed against the *whole*
 * reconstructed file (its own 4 bytes treated as zero while computing that
 * checksum, then overwritten with the real adjustment) — the standard sfnt
 * checksum convention, needed because a stricter font-loading path than
 * "just grab the glyphs" (some PDF viewers, most real OS font loaders) can
 * reject a font whose checksums don't add up. Shared by both the WOFF1
 * decoder above and the WOFF2 decoder (woff2-decoder.js), which is why it's
 * exported rather than kept private to this file.
 */
export function buildSfnt(flavor, tables) {
  const sorted = [...tables].sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0));
  const numTables = sorted.length;

  let entrySelector = 0;
  while (2 ** (entrySelector + 1) <= numTables) entrySelector++;
  const searchRange = 2 ** entrySelector * 16;
  const rangeShift = numTables * 16 - searchRange;

  const headerAndDirLength = 12 + numTables * 16;
  let bodyOffset = headerAndDirLength;
  const layout = sorted.map(({ tag, data }) => {
    const offset = bodyOffset;
    const padded = (data.length + 3) & ~3;
    bodyOffset += padded;
    return { tag, data, offset };
  });

  const w = new ByteWriter();
  w.writeU32be(flavor);
  w.writeU16be(numTables);
  w.writeU16be(searchRange);
  w.writeU16be(entrySelector);
  w.writeU16be(rangeShift);
  for (const { tag, data, offset } of layout) {
    w.writeString(tag, 'latin1');
    w.writeU32be(tableChecksum(data));
    w.writeU32be(offset);
    w.writeU32be(data.length);
  }
  for (const { data } of layout) {
    w.writeBytes(data);
    const pad = (4 - (data.length % 4)) % 4;
    if (pad) w.writeBytes(Buffer.alloc(pad));
  }

  const sfnt = w.toBuffer();

  const headEntry = layout.find((t) => t.tag === 'head');
  if (headEntry) {
    // Zero checksumAdjustment (head offset +8) before computing the whole-
    // file checksum, per the sfnt spec's own bootstrapping rule — otherwise
    // the field would be checksummed against its own prior (wrong) value.
    sfnt.writeUInt32BE(0, headEntry.offset + 8);
    const wholeFileChecksum = tableChecksum(sfnt);
    const adjustment = (0xb1b0afba - wholeFileChecksum) >>> 0;
    sfnt.writeUInt32BE(adjustment, headEntry.offset + 8);
  }

  return sfnt;
}
