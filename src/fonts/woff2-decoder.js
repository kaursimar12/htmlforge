/**
 * WOFF2 container decoder. Unlike WOFF1 (woff-decoder.js), WOFF2 Brotli-
 * compresses every table's data together as ONE stream (not per-table), and
 * — critically — most WOFF2 encoders (including every mainstream Google
 * Fonts/Adobe Fonts/etc. pipeline) apply a special reversible TRANSFORM to
 * `glyf`/`loca` before compression: a custom, more compact re-encoding of
 * glyph outlines that is NOT valid sfnt `glyf`/`loca` data as-is and must be
 * reconstructed back into real outline data before the font is usable. That
 * reconstruction is implemented below (reconstructGlyf) — ported field-for-
 * field from fontTools' own `WOFF2GlyfTable.reconstruct`/`_decodeTriplets`
 * (MIT-licensed reference implementation already used by essentially every
 * real-world font tool) rather than re-derived from the spec text alone,
 * since the point/flag "triplet" bit-packing is fiddly enough that a
 * subtly-wrong reimplementation would silently corrupt glyph shapes instead
 * of throwing. Verified byte-for-byte against fontTools' own decode of real
 * downloaded fonts in tests/woff2-transform.test.js.
 *
 * `hmtx` has its own separate, optional WOFF2 transform (unrelated to the
 * glyf/loca one) that this decoder does NOT implement yet — still throws a
 * specific error for that case rather than guessing.
 */
import { brotliDecompressSync } from 'node:zlib';
import { ByteReader, ByteWriter } from '../utils/bytes.js';
import { buildSfnt } from './woff-decoder.js';

const WOFF2_SIGNATURE = 0x774f4632; // 'wOF2'

const KNOWN_TABLE_TAGS = [
  'cmap', 'head', 'hhea', 'hmtx', 'maxp', 'name', 'OS/2', 'post',
  'cvt ', 'fpgm', 'glyf', 'loca', 'prep', 'CFF ', 'VORG', 'EBDT',
  'EBLC', 'gasp', 'hdmx', 'kern', 'LTSH', 'PCLT', 'VDMX', 'vhea',
  'vmtx', 'BASE', 'GDEF', 'GPOS', 'GSUB', 'EBSC', 'JSTF', 'MATH',
  'CBDT', 'CBLC', 'COLR', 'CPAL', 'SVG ', 'sbix', 'acnt', 'avar',
  'bdat', 'bloc', 'bsln', 'cvar', 'fdsc', 'feat', 'fmtx', 'fvar',
  'gvar', 'hsty', 'just', 'lcar', 'mort', 'morx', 'opbd', 'prop',
  'trak', 'Zapf', 'Silf', 'Glat', 'Gloc', 'Feat', 'Sill',
];

export function isWoff2(buffer) {
  return buffer.length >= 4 && buffer.readUInt32BE(0) === WOFF2_SIGNATURE;
}

// WOFF2's UIntBase128: a big-endian base-128 varint — each byte's top bit
// marks "more bytes follow"; capped at 5 bytes since that's the max needed
// to hold a 32-bit value in 7-bit groups, guarding against a corrupt/hostile
// file looping forever with an unterminated high-bit run.
function readUIntBase128(r) {
  let value = 0;
  for (let i = 0; i < 5; i++) {
    const b = r.u8();
    value = value * 128 + (b & 0x7f);
    if ((b & 0x80) === 0) return value;
  }
  throw new Error('Malformed UIntBase128 (too many continuation bytes)');
}

// WOFF2's 255UInt16: a different, separate compact varint used inside the
// transformed-glyf sub-streams (contour point counts, instruction lengths) —
// values below 253 are one byte; 253/254/255 are escape codes for a
// following byte (see the WOFF2 spec's own worked examples, matched exactly
// by fontTools' unpack255UShort, which this ports).
function readPacked255UShort(r) {
  const code = r.u8();
  if (code < 253) return code;
  if (code === 253) return r.u16be();
  if (code === 254) return 506 + r.u8();
  return 253 + r.u8(); // code === 255
}

function readTableDirectory(r, numTables) {
  const entries = [];
  for (let i = 0; i < numTables; i++) {
    const flags = r.u8();
    const knownIndex = flags & 0x3f;
    const tag = knownIndex === 0x3f ? r.string(4, 'latin1') : KNOWN_TABLE_TAGS[knownIndex];
    if (!tag) throw new Error(`Unknown WOFF2 table index ${knownIndex}`);
    const xformVersion = (flags >> 6) & 0x3;
    const origLength = readUIntBase128(r);
    // Per the WOFF2 spec: for glyf/loca, xform version 0 (the default) means
    // the transform WAS applied; version 3 means it wasn't. For every other
    // table, version 0 means NOT transformed — except hmtx, which a later
    // spec amendment lets encoders optionally transform via version 1. Any
    // other combination we don't recognize is treated as "transformed" too,
    // the safe direction to err on (bail out rather than assume raw bytes).
    const isGlyfOrLoca = tag === 'glyf' || tag === 'loca';
    const transformed = isGlyfOrLoca ? xformVersion !== 3 : tag === 'hmtx' ? xformVersion === 1 : xformVersion !== 0;
    const transformLength = transformed ? readUIntBase128(r) : undefined;
    entries.push({ tag, origLength, transformLength, transformed });
  }
  return entries;
}

// Standard sfnt simple-glyph flag bits this decoder ever needs to set —
// every other bit (X_SHORT/Y_SHORT/REPEAT/SAME_OR_POSITIVE) is deliberately
// left at 0 by always emitting an explicit signed 16-bit delta per
// coordinate below: a larger `glyf` table than a size-optimized encoder
// would produce, but unambiguously spec-valid (a 0 delta is a perfectly
// legal explicit value, not something that must use the "same" shortcut),
// which matters far more than a few extra bytes for an embedded PDF font.
const FLAG_ON_CURVE = 0x01;
const FLAG_OVERLAP_SIMPLE = 0x40;

// Composite glyph component flags (standard sfnt `glyf` format — WOFF2's
// compositeStream stores these components pre-encoded exactly as they'd
// appear in a normal glyf table, per the WOFF2 spec, so they're copied
// through field-by-field rather than reinterpreted).
const ARG_1_AND_2_ARE_WORDS = 0x0001;
const WE_HAVE_A_SCALE = 0x0008;
const MORE_COMPONENTS = 0x0020;
const WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
const WE_HAVE_A_TWO_BY_TWO = 0x0080;
const WE_HAVE_INSTRUCTIONS = 0x0100;

/** `flag`'s LSB selects the sign of `magnitude` (1 = positive) — matches
 * fontTools' own `withSign` helper exactly (the same "which bit means what"
 * this whole transform is fiddly enough to get wrong from the spec alone). */
function withSign(signBit, magnitude) {
  return (signBit & 1) ? magnitude : -magnitude;
}

/**
 * Decodes one point's (dx, dy) delta from WOFF2's "triplet" coordinate
 * encoding — `flag` is the point's format code (0-127, already stripped of
 * its on-curve bit), and `glyphStream` is positioned at that format's own
 * extra bytes. Ported line-for-line from fontTools' `_decodeTriplets` inner
 * loop (see this file's own top comment for why this isn't re-derived from
 * the spec text alone).
 */
function readTriplet(flag, glyphStream) {
  if (flag < 10) {
    const b0 = glyphStream.u8();
    return [0, withSign(flag, ((flag & 14) << 7) + b0)];
  }
  if (flag < 20) {
    const b0 = glyphStream.u8();
    return [withSign(flag, (((flag - 10) & 14) << 7) + b0), 0];
  }
  if (flag < 84) {
    const b0v = flag - 20;
    const b1 = glyphStream.u8();
    const dx = withSign(flag, 1 + (b0v & 0x30) + (b1 >> 4));
    const dy = withSign(flag >> 1, 1 + ((b0v & 0x0c) << 2) + (b1 & 0x0f));
    return [dx, dy];
  }
  if (flag < 120) {
    const b0v = flag - 84;
    const b1 = glyphStream.u8();
    const b2 = glyphStream.u8();
    const dx = withSign(flag, 1 + (Math.floor(b0v / 12) << 8) + b1);
    const dy = withSign(flag >> 1, 1 + (((b0v % 12) >> 2) << 8) + b2);
    return [dx, dy];
  }
  if (flag < 124) {
    const b1 = glyphStream.u8();
    const b2 = glyphStream.u8();
    const b3 = glyphStream.u8();
    const dx = withSign(flag, (b1 << 4) + (b2 >> 4));
    const dy = withSign(flag >> 1, ((b2 & 0x0f) << 8) + b3);
    return [dx, dy];
  }
  const b1 = glyphStream.u8();
  const b2 = glyphStream.u8();
  const b3 = glyphStream.u8();
  const b4 = glyphStream.u8();
  const dx = withSign(flag, (b1 << 8) + b2);
  const dy = withSign(flag >> 1, (b3 << 8) + b4);
  return [dx, dy];
}

/**
 * Reconstructs real `glyf`/`loca` table bytes from WOFF2's transformed glyf
 * representation (`data` is the transformed table's own bytes, sliced by
 * `transformLength` — NOT `origLength` — from the decompressed blob). The
 * transform stores every glyph's outline across several separately-batched
 * sub-streams (contour counts, per-contour point counts, per-point on-curve
 * flags, delta-encoded "triplet" coordinates, composite component records,
 * bounding boxes, hinting instructions) instead of one independent record
 * per glyph — this walks all of them back into standard per-glyph sfnt
 * records, one glyph at a time, in glyph-ID order (the sub-streams are only
 * meaningful read sequentially in that order; there's no per-glyph offset
 * table to jump to one directly).
 */
function reconstructGlyf(data) {
  const r = new ByteReader(data);

  const version = r.u16be();
  const optionFlags = r.u16be();
  const numGlyphs = r.u16be();
  const indexFormat = r.u16be();
  const nContourStreamSize = r.u32be();
  const nPointsStreamSize = r.u32be();
  const flagStreamSize = r.u32be();
  const glyphStreamSize = r.u32be();
  const compositeStreamSize = r.u32be();
  const bboxStreamSize = r.u32be();
  const instructionStreamSize = r.u32be();
  void version;

  const nContourStream = new ByteReader(r.bytes(nContourStreamSize));
  const nPointsStream = new ByteReader(r.bytes(nPointsStreamSize));
  const flagStream = r.bytes(flagStreamSize); // one raw byte per point — read positionally, no varint decoding
  const glyphStream = new ByteReader(r.bytes(glyphStreamSize));
  const compositeStream = new ByteReader(r.bytes(compositeStreamSize));
  const bboxStreamAll = r.bytes(bboxStreamSize);
  const instructionStream = new ByteReader(r.bytes(instructionStreamSize));

  // Bit 0 of optionFlags: an extra per-glyph bitmap follows every declared
  // sub-stream, flagging which simple glyphs had their first point's
  // OVERLAP_SIMPLE bit set (an outline-rendering hint, not shape data).
  const hasOverlapSimpleBitmap = (optionFlags & 0x0001) !== 0;
  const overlapSimpleBitmap = hasOverlapSimpleBitmap ? r.bytes((numGlyphs + 7) >> 3) : undefined;

  if (!r.eof) {
    throw new Error(`WOFF2 transformed glyf table has ${r.remaining} unexpected trailing byte(s)`);
  }

  // bboxStream itself opens with a per-glyph "does this glyph have an
  // explicit bbox" bitmap (rounded up to a 4-byte/32-glyph boundary),
  // followed by the actual 4×Int16 bbox records for whichever glyphs it
  // flagged — composite glyphs always have one (there's no cheap way to
  // compute a composite's bbox without resolving every referenced
  // component), simple glyphs only when the encoder chose not to recompute
  // it from the decoded points itself.
  const bboxBitmapSize = ((numGlyphs + 31) >> 5) << 2;
  const bboxBitmap = bboxStreamAll.subarray(0, bboxBitmapSize);
  const bboxStream = new ByteReader(bboxStreamAll.subarray(bboxBitmapSize));

  const hasBBoxBit = (glyphId) => Boolean(bboxBitmap[glyphId >> 3] & (0x80 >> (glyphId & 7)));
  const hasOverlapBit = (glyphId) => hasOverlapSimpleBitmap && Boolean(overlapSimpleBitmap[glyphId >> 3] & (0x80 >> (glyphId & 7)));

  let flagCursor = 0;
  const glyphBuffers = [];

  for (let glyphId = 0; glyphId < numGlyphs; glyphId++) {
    const nContours = nContourStream.i16be();

    if (nContours === 0) {
      glyphBuffers.push(Buffer.alloc(0)); // empty glyph — e.g. space — no data, no bbox
      continue;
    }

    const w = new ByteWriter();

    if (nContours > 0) {
      // --- simple glyph -----------------------------------------------
      const endPtsOfContours = [];
      let endPoint = -1;
      for (let c = 0; c < nContours; c++) {
        endPoint += readPacked255UShort(nPointsStream);
        endPtsOfContours.push(endPoint);
      }
      const nPoints = endPtsOfContours[endPtsOfContours.length - 1] + 1;

      const onCurve = new Array(nPoints);
      const xs = new Array(nPoints);
      const ys = new Array(nPoints);
      let x = 0;
      let y = 0;
      for (let p = 0; p < nPoints; p++) {
        const rawFlag = flagStream[flagCursor++];
        const isOnCurve = (rawFlag >> 7) === 0;
        const [dx, dy] = readTriplet(rawFlag & 0x7f, glyphStream);
        x += dx;
        y += dy;
        onCurve[p] = isOnCurve;
        xs[p] = x;
        ys[p] = y;
      }

      const overlapSimple = hasOverlapBit(glyphId);
      const instrLength = readPacked255UShort(glyphStream);
      const instructions = instructionStream.bytes(instrLength);

      let xMin;
      let yMin;
      let xMax;
      let yMax;
      if (hasBBoxBit(glyphId)) {
        xMin = bboxStream.i16be(); yMin = bboxStream.i16be(); xMax = bboxStream.i16be(); yMax = bboxStream.i16be();
      } else {
        xMin = Math.min(...xs); xMax = Math.max(...xs);
        yMin = Math.min(...ys); yMax = Math.max(...ys);
      }

      w.writeU16be(nContours);
      w.writeU16be(xMin & 0xffff); w.writeU16be(yMin & 0xffff); w.writeU16be(xMax & 0xffff); w.writeU16be(yMax & 0xffff);
      for (const e of endPtsOfContours) w.writeU16be(e);
      w.writeU16be(instrLength);
      w.writeBytes(instructions);
      for (let p = 0; p < nPoints; p++) {
        let flagByte = onCurve[p] ? FLAG_ON_CURVE : 0;
        if (p === 0 && overlapSimple) flagByte |= FLAG_OVERLAP_SIMPLE;
        w.writeU8(flagByte);
      }
      // Standard glyf coordinates are stored as two SEPARATE contiguous
      // blocks — every point's x delta first, then every point's y delta —
      // not interleaved (x0,y0,x1,y1,...); each delta is relative to the
      // previous point (first point's delta is from the origin), not the
      // absolute position `xs`/`ys` hold (WOFF2's own reconstructed running
      // total, kept above for bbox/min/max computation).
      for (let p = 0; p < nPoints; p++) {
        w.writeU16be((xs[p] - (p === 0 ? 0 : xs[p - 1])) & 0xffff);
      }
      for (let p = 0; p < nPoints; p++) {
        w.writeU16be((ys[p] - (p === 0 ? 0 : ys[p - 1])) & 0xffff);
      }
    } else {
      // --- composite glyph ---------------------------------------------
      // Copied through field-by-field (not reinterpreted) — the WOFF2 spec
      // stores these components pre-encoded exactly as a normal glyf table
      // would, so only the flag bits are read here at all, purely to know
      // how many bytes each field occupies.
      const compWriter = new ByteWriter();
      let more = true;
      let haveInstructions = false;
      while (more) {
        const flags = compositeStream.u16be();
        const glyphIndex = compositeStream.u16be();
        compWriter.writeU16be(flags);
        compWriter.writeU16be(glyphIndex);
        compWriter.writeBytes(compositeStream.bytes(flags & ARG_1_AND_2_ARE_WORDS ? 4 : 2));
        const scaleBytes = flags & WE_HAVE_A_SCALE ? 2 : flags & WE_HAVE_AN_X_AND_Y_SCALE ? 4 : flags & WE_HAVE_A_TWO_BY_TWO ? 8 : 0;
        if (scaleBytes) compWriter.writeBytes(compositeStream.bytes(scaleBytes));
        if (flags & WE_HAVE_INSTRUCTIONS) haveInstructions = true;
        more = Boolean(flags & MORE_COMPONENTS);
      }

      let instrLength = 0;
      let instructions;
      if (haveInstructions) {
        instrLength = readPacked255UShort(glyphStream);
        instructions = instructionStream.bytes(instrLength);
      }

      if (!hasBBoxBit(glyphId)) throw new Error(`WOFF2 composite glyph ${glyphId} has no bbox in bboxStream (required by spec)`);
      const xMin = bboxStream.i16be();
      const yMin = bboxStream.i16be();
      const xMax = bboxStream.i16be();
      const yMax = bboxStream.i16be();

      w.writeU16be(nContours & 0xffff); // negative — written as its 16-bit two's complement
      w.writeU16be(xMin & 0xffff); w.writeU16be(yMin & 0xffff); w.writeU16be(xMax & 0xffff); w.writeU16be(yMax & 0xffff);
      w.writeBytes(compWriter.toBuffer());
      if (haveInstructions) {
        w.writeU16be(instrLength);
        w.writeBytes(instructions);
      }
    }

    glyphBuffers.push(w.toBuffer());
  }

  // Every glyph must start at an even offset — pad any odd-length record by
  // one zero byte (never needed for empty glyphs, already length 0).
  const paddedGlyphs = glyphBuffers.map((b) => (b.length % 2 === 0 ? b : Buffer.concat([b, Buffer.alloc(1)])));
  const glyf = Buffer.concat(paddedGlyphs);

  const locaWriter = new ByteWriter();
  let running = 0;
  locaWriter[indexFormat === 0 ? 'writeU16be' : 'writeU32be'](indexFormat === 0 ? running / 2 : running);
  for (const g of paddedGlyphs) {
    running += g.length;
    locaWriter[indexFormat === 0 ? 'writeU16be' : 'writeU32be'](indexFormat === 0 ? running / 2 : running);
  }

  return { glyf, loca: locaWriter.toBuffer(), indexFormat };
}

/**
 * Decodes a WOFF2 buffer into a standalone sfnt buffer (same contract as
 * woff-decoder.js's decodeWoff). Throws — for the caller to warn on and fall
 * back, never to crash the export — on a malformed file, on the (separate,
 * unrelated) `hmtx` transform this decoder doesn't reconstruct, or on any
 * other table transform this file doesn't recognize.
 */
export function decodeWoff2(buffer) {
  const r = new ByteReader(buffer);
  const signature = r.u32be();
  if (signature !== WOFF2_SIGNATURE) throw new Error('Not a WOFF2 file (bad signature)');
  const flavor = r.u32be();
  r.skip(4); // length
  const numTables = r.u16be();
  r.skip(2); // reserved
  r.skip(4); // totalSfntSize
  const totalCompressedSize = r.u32be();
  r.skip(2 + 2); // majorVersion, minorVersion
  r.skip(4 + 4 + 4); // metaOffset, metaLength, metaOrigLength
  r.skip(4 + 4); // privOffset, privLength

  const directory = readTableDirectory(r, numTables);

  // glyf/loca are reconstructed below; hmtx has its own separate transform
  // this decoder doesn't implement yet; anything else coming back
  // "transformed" is unexpected — safer to bail out than guess.
  for (const t of directory) {
    if (t.transformed && t.tag === 'hmtx') {
      throw new Error(`Font uses WOFF2's "hmtx" transform, which this decoder doesn't reconstruct yet — cannot safely embed`);
    }
    if (t.transformed && t.tag !== 'glyf' && t.tag !== 'loca' && t.tag !== 'hmtx') {
      throw new Error(`Font uses WOFF2's "${t.tag}" transform, which this decoder doesn't reconstruct — cannot safely embed`);
    }
  }

  const compressedStart = r.offset;
  const compressed = buffer.subarray(compressedStart, compressedStart + totalCompressedSize);
  const decompressed = brotliDecompressSync(compressed);

  let cursor = 0;
  const tables = directory.map(({ tag, origLength, transformLength, transformed }) => {
    const size = transformed ? transformLength : origLength;
    const data = decompressed.subarray(cursor, cursor + size);
    cursor += size;
    return { tag, data, transformed };
  });

  const glyfEntry = tables.find((t) => t.tag === 'glyf');
  if (glyfEntry?.transformed) {
    const { glyf, loca, indexFormat } = reconstructGlyf(glyfEntry.data);
    glyfEntry.data = glyf;
    const locaEntry = tables.find((t) => t.tag === 'loca');
    if (locaEntry) locaEntry.data = loca;
    const headEntry = tables.find((t) => t.tag === 'head');
    if (headEntry) {
      const patched = Buffer.from(headEntry.data);
      patched.writeUInt16BE(indexFormat, 50); // head.indexToLocFormat
      headEntry.data = patched;
    }
  }

  return buildSfnt(flavor, tables);
}
