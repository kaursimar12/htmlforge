/**
 * Best-effort intrinsic pixel-dimension reading from raw image bytes — no
 * npm dependency, no full pixel decode (contrast exporter/pdf/images.js's
 * own PNG decoder, which unfilters every scanline because it needs the real
 * pixel data for a PDF XObject; this only ever reads a handful of header
 * bytes). Used by the asset resolver (assets/resolver.js) to stamp a
 * resolved image's own real width/height once, so DOCX/PPTX/XLSX can fall
 * back to it — instead of a flat placeholder size — whenever the source
 * HTML declared no `width`/`height` of its own.
 *
 * `parseJpegSof` is the same JPEG SOF-marker scan exporter/pdf/images.js
 * needs anyway for its own DCTDecode XObject sizing/colorspace — shared
 * verbatim so there's exactly one JPEG-marker-walking implementation.
 */
import { ByteReader } from './bytes.js';

const SOF_MARKERS = new Set([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf]);

/** `{ width, height, components, precision }` or `undefined` if this isn't a real JPEG. */
export function parseJpegSof(buffer) {
  const r = new ByteReader(buffer);
  if (r.u16be() !== 0xffd8) return undefined; // not a real JPEG SOI marker
  while (!r.eof) {
    const marker = r.u16be();
    if ((marker & 0xff00) !== 0xff00) return undefined; // desynced, bail out
    const kind = marker & 0xff;
    if (kind === 0xd8 || kind === 0xd9 || (kind >= 0xd0 && kind <= 0xd7)) continue; // no length field
    const segmentLength = r.u16be();
    if (SOF_MARKERS.has(kind)) {
      const precision = r.u8();
      const height = r.u16be();
      const width = r.u16be();
      const components = r.u8();
      return { width, height, components, precision };
    }
    r.skip(segmentLength - 2);
  }
  return undefined;
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

/**
 * `{ width, height }` or `undefined` — reads only the PNG signature + IHDR
 * chunk header (the first chunk in any valid PNG), never the IDAT pixel
 * data — IHDR's own width/height fields are already real pixel counts
 * regardless of bit depth/color type/interlacing, so nothing else in the
 * file needs to be read just to answer "how big is this image."
 */
export function readPngDimensions(buffer) {
  const r = new ByteReader(buffer);
  for (const byte of PNG_SIGNATURE) if (r.u8() !== byte) return undefined;
  r.skip(4); // IHDR chunk length (always 13, not needed)
  if (r.string(4, 'ascii') !== 'IHDR') return undefined;
  const width = r.u32be();
  const height = r.u32be();
  return { width, height };
}

/** `{ width, height }` or `undefined` — dispatches by the asset's own already-detected mimeType. */
export function readImageDimensions(buffer, mimeType) {
  try {
    if (mimeType === 'image/jpeg') {
      const sof = parseJpegSof(buffer);
      return sof && { width: sof.width, height: sof.height };
    }
    if (mimeType === 'image/png') return readPngDimensions(buffer);
  } catch {
    // A truncated/corrupt image can make these bounds-checked reads throw
    // mid-parse — dimensions are a nice-to-have fallback, never worth
    // failing asset resolution over (same recoverable-vs-fatal split
    // CLAUDE.md's error-handling contract already uses everywhere else).
    return undefined;
  }
  return undefined;
}
