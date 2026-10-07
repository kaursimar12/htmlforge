/**
 * Image → PDF XObject conversion.
 *
 * - JPEG: always a direct `DCTDecode` passthrough — the raw compressed
 *   bytes go straight into the XObject stream unmodified. This works for
 *   both baseline and progressive JPEG, since `DCTDecode` decodes the same
 *   JPEG bitstream a browser would.
 * - PNG: a real decoder — unfilters every scanline (PNG's own five adaptive
 *   per-row filters: None/Sub/Up/Average/Paeth) and expands palette
 *   (colorType 3) indices via `PLTE` — covering the color types that
 *   actually show up in practice: grayscale, truecolor, palette, and both
 *   of their alpha-carrying forms (grayscale+alpha, RGBA), the last two of
 *   which are what almost every real screenshot/logo/placeholder-service
 *   image actually is. An alpha channel becomes a real PDF `/SMask` (a
 *   sibling grayscale XObject referenced from the color image's own dict),
 *   not just dropped — a transparent logo should still look transparent.
 *   16-bit-per-channel and interlaced (Adam7) PNGs remain a documented,
 *   recoverable skip; those are rare enough on the web that decoding them
 *   isn't worth the added complexity in this pass.
 *
 * GIF/WebP/BMP (and any other format `resolveAssets`/`magic.js` might
 * sniff) are not embedded at all — same recoverable skip.
 */
import zlib from 'node:zlib';
import { ByteReader } from '../../utils/bytes.js';
import { parseJpegSof } from '../../utils/image-dimensions.js';
import { name } from './pdf-writer.js';

// parseJpegSof (utils/image-dimensions.js) scans a JPEG's marker segments
// for the Start-Of-Frame marker, which carries the real pixel dimensions and
// component count — PDF's own `/Width`/`/Height`/`/ColorSpace` on the
// XObject dict need these regardless of the "no real decode" passthrough
// approach below, since DCTDecode doesn't derive them from the stream
// itself the way a full JPEG decoder would. Shared with the asset resolver's
// own (lighter-weight, decode-free) dimension reading — see that module's
// own doc comment for why this exact function is the one place that walks
// JPEG marker segments.

function jpegXObject(data, resourceName) {
  const sof = parseJpegSof(data);
  if (!sof) return undefined;
  // A DCTDecode stream's own component count determines PDF's own
  // colorSpace: 1 = a grayscale JPEG, 3 = the standard YCbCr-as-RGB case
  // every ordinary photo/web JPEG uses. 4-component (CMYK/YCCK) JPEGs are
  // uncommon on the web and need an /Adobe-marker-aware transform decision
  // this pass doesn't make — treated as unsupported, same as an
  // unrecognized PNG variant.
  const colorSpace = sof.components === 1 ? 'DeviceGray' : sof.components === 3 ? 'DeviceRGB' : undefined;
  if (!colorSpace) return undefined;
  return {
    resourceName,
    data,
    dict: {
      Type: name('XObject'),
      Subtype: name('Image'),
      Width: sof.width,
      Height: sof.height,
      ColorSpace: name(colorSpace),
      BitsPerComponent: 8,
      Filter: name('DCTDecode'),
    },
  };
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function readPngChunks(buffer) {
  const r = new ByteReader(buffer);
  const sig = r.bytes(8);
  for (let i = 0; i < 8; i++) if (sig[i] !== PNG_SIGNATURE[i]) return undefined;

  const chunks = [];
  while (!r.eof) {
    const length = r.u32be();
    // A chunk declaring more data than actually remains in the buffer means
    // this file is truncated or corrupt (a network fetch cut off
    // mid-download is the common real-world cause) — bail out cleanly
    // rather than letting the subsequent `r.bytes(length)` throw a raw
    // ByteReader bounds error past this function's own control (still
    // caught one level up in buildImageXObject either way, but this gives a
    // clearer, PNG-specific signal before that safety net has to catch it).
    if (length > r.remaining - 4) return undefined;
    const type = r.string(4, 'ascii');
    const data = r.bytes(length);
    r.skip(4); // CRC, not verified — this is a decoder, not a PNG validator
    chunks.push({ type, data });
    if (type === 'IEND') break;
  }
  return chunks;
}

// The number of channels PNG itself packs per pixel for each color type —
// what the file's own scanlines are made of, before this decoder expands
// palette indices or splits out an alpha channel.
const FILE_CHANNELS_BY_COLOR_TYPE = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

/**
 * Undoes PNG's own per-scanline adaptive filtering (spec §9) — each row is
 * independently filtered against the row above it and the pixel to its own
 * left, so reconstruction must proceed top-to-bottom, left-to-right, one
 * row's own *already-reconstructed* bytes feeding the next. `bpp` (bytes per
 * complete pixel) is what the spec's own Paeth/Sub/Average filters use as
 * "distance to the pixel on the left" — at 8 bits/channel this is simply the
 * channel count, never a fractional byte the way <8-bit depths would need.
 */
function unfilterScanlines(inflated, width, height, channels) {
  const rowBytes = width * channels;
  const bpp = channels;
  const out = Buffer.alloc(rowBytes * height);
  let prevRow = Buffer.alloc(rowBytes); // an implicit all-zero row above row 0
  let pos = 0;
  for (let y = 0; y < height; y++) {
    const filterType = inflated[pos];
    pos += 1;
    const row = inflated.subarray(pos, pos + rowBytes);
    pos += rowBytes;
    const outRow = out.subarray(y * rowBytes, y * rowBytes + rowBytes);
    for (let x = 0; x < rowBytes; x++) {
      const a = x >= bpp ? outRow[x - bpp] : 0; // reconstructed byte to the left
      const b = prevRow[x]; // reconstructed byte directly above
      const c = x >= bpp ? prevRow[x - bpp] : 0; // reconstructed byte above-left
      let val = row[x];
      if (filterType === 1) val = (val + a) & 0xff; // Sub
      else if (filterType === 2) val = (val + b) & 0xff; // Up
      else if (filterType === 3) val = (val + ((a + b) >> 1)) & 0xff; // Average
      else if (filterType === 4) { // Paeth
        const p = a + b - c;
        const pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
        val = (val + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c)) & 0xff;
      } else if (filterType !== 0) {
        throw new Error(`unrecognized PNG scanline filter type ${filterType}`);
      }
      outRow[x] = val;
    }
    prevRow = outRow;
  }
  return out;
}

/**
 * Splits a PNG's own unfiltered pixel data into a plain color buffer (1
 * channel for grayscale, 3 for RGB — what PDF's DeviceGray/DeviceRGB expect)
 * and, when the source carried one, a separate 1-byte-per-pixel alpha
 * buffer for use as a `/SMask`. Palette (colorType 3) expansion looks up
 * each index byte in `palette` (PLTE's own raw RGB-triplet bytes) and, when
 * the PNG declared a `tRNS` chunk, that index's own alpha value.
 */
function splitColorAndAlpha(pixels, width, height, colorType, palette, trns) {
  const count = width * height;
  if (colorType === 0 || colorType === 2) return { channels: colorType === 0 ? 1 : 3, color: pixels, alpha: undefined };

  if (colorType === 4) { // grayscale + alpha: 2 bytes/pixel
    const color = Buffer.alloc(count);
    const alpha = Buffer.alloc(count);
    for (let i = 0; i < count; i++) { color[i] = pixels[i * 2]; alpha[i] = pixels[i * 2 + 1]; }
    return { channels: 1, color, alpha };
  }
  if (colorType === 6) { // truecolor + alpha: 4 bytes/pixel
    const color = Buffer.alloc(count * 3);
    const alpha = Buffer.alloc(count);
    for (let i = 0; i < count; i++) {
      color[i * 3] = pixels[i * 4];
      color[i * 3 + 1] = pixels[i * 4 + 1];
      color[i * 3 + 2] = pixels[i * 4 + 2];
      alpha[i] = pixels[i * 4 + 3];
    }
    return { channels: 3, color, alpha };
  }
  // colorType === 3: palette — one index byte per pixel.
  const color = Buffer.alloc(count * 3);
  const alpha = trns ? Buffer.alloc(count) : undefined;
  for (let i = 0; i < count; i++) {
    const idx = pixels[i];
    color[i * 3] = palette[idx * 3] ?? 0;
    color[i * 3 + 1] = palette[idx * 3 + 1] ?? 0;
    color[i * 3 + 2] = palette[idx * 3 + 2] ?? 0;
    if (alpha) alpha[i] = idx < trns.length ? trns[idx] : 255;
  }
  return { channels: 3, color, alpha };
}

function imageDict(width, height, channels) {
  return {
    Type: name('XObject'),
    Subtype: name('Image'),
    Width: width,
    Height: height,
    ColorSpace: name(channels === 1 ? 'DeviceGray' : 'DeviceRGB'),
    BitsPerComponent: 8,
    Filter: name('FlateDecode'),
  };
}

function pngXObject(data, resourceName, warn) {
  const chunks = readPngChunks(data);
  if (!chunks) return undefined;
  const ihdr = chunks.find((c) => c.type === 'IHDR');
  if (!ihdr) return undefined;

  const r = new ByteReader(ihdr.data);
  const width = r.u32be();
  const height = r.u32be();
  const bitDepth = r.u8();
  const colorType = r.u8();
  r.u8(); // compression method — always 0 for PNG, not needed
  r.u8(); // filter method — always 0 for PNG, not needed
  const interlaceMethod = r.u8();

  const fileChannels = FILE_CHANNELS_BY_COLOR_TYPE[colorType];
  const supported = bitDepth === 8 && interlaceMethod === 0 && fileChannels !== undefined;
  if (!supported) {
    warn?.('pdf-unsupported-png', `Skipping PNG (colorType=${colorType}, bitDepth=${bitDepth}, interlace=${interlaceMethod}): only non-interlaced 8-bit PNGs (grayscale, RGB, palette, or either with alpha) are supported.`);
    return undefined;
  }

  let palette, trns;
  if (colorType === 3) {
    const plte = chunks.find((c) => c.type === 'PLTE');
    if (!plte) { warn?.('pdf-unsupported-png', 'Skipping palette PNG with no PLTE chunk.'); return undefined; }
    palette = plte.data;
    trns = chunks.find((c) => c.type === 'tRNS')?.data;
  }

  const idat = Buffer.concat(chunks.filter((c) => c.type === 'IDAT').map((c) => c.data));
  const inflated = zlib.inflateSync(idat);
  const pixels = unfilterScanlines(inflated, width, height, fileChannels);
  const { channels, color, alpha } = splitColorAndAlpha(pixels, width, height, colorType, palette, trns);

  let smask;
  if (alpha) {
    smask = {
      resourceName: `${resourceName}Smask`,
      data: zlib.deflateSync(alpha),
      dict: imageDict(width, height, 1),
    };
  }

  return {
    resourceName,
    data: zlib.deflateSync(color),
    dict: { ...imageDict(width, height, channels), SMask: smask ? true : undefined },
    smask,
  };
}

/** Returns `{resourceName, data, dict, smask?}` or `undefined` (skip, already warned) for one resolved asset. */
export function buildImageXObject(asset, resourceName, warn) {
  if (!asset.data) return undefined;
  try {
    if (asset.mimeType === 'image/jpeg') return jpegXObject(asset.data, resourceName);
    if (asset.mimeType === 'image/png') return pngXObject(asset.data, resourceName, warn);
  } catch (err) {
    // A truncated/corrupted fetch — or any other malformed byte stream that
    // merely LOOKS like a JPEG/PNG (a valid signature followed by garbage
    // or cut-off data) — can make this parser's own bounds-checked reads
    // (ByteReader's own #ensure, or zlib's own inflate) throw mid-parse.
    // Per this project's own error-handling contract (CLAUDE.md's
    // Recoverable/Fatal split), a bad IMAGE is explicitly a "skip it, warn,
    // keep going" case, never fatal — without this catch, one corrupt
    // image previously crashed the ENTIRE document's conversion, not just
    // that one image.
    warn?.('pdf-corrupt-image', `Skipping image with unparseable/corrupt data for src "${asset.src}": ${err.message}`);
    return undefined;
  }
  warn?.('pdf-unsupported-image-format', `Skipping image with unsupported format "${asset.mimeType}" for src "${asset.src}" — only JPEG and PNG can be embedded.`);
  return undefined;
}
