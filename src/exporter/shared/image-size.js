/**
 * Resolves an image's display size in CSS px — one shared "real beats
 * declared beats placeholder" precedence used identically by DOCX/PPTX/XLSX
 * (each converts the result to its own unit — EMU for DOCX/PPTX, px directly
 * for XLSX — and applies its own max-width clamp):
 *
 *   1. the image's own real intrinsic pixel dimensions, read from its actual
 *      bytes once during asset resolution (assets/resolver.js's own
 *      width/height fields — see utils/image-dimensions.js) — a real
 *      measurement always wins, even over an explicit HTML/CSS width or
 *      height, since the source markup's declared display size is routinely
 *      wrong/stale relative to the actual file (a placeholder-era `<img
 *      style="width:80px;height:80px">` pointing at a since-swapped full-res
 *      photo) and this project's own output should reflect what the image
 *      actually is, not what the page claimed it was.
 *   2. the source HTML's own explicit `width`/`height` (ImageNode.width/
 *      height) — only a fallback for a format this project can't decode
 *      (no npm dependency means GIF/WebP/BMP/SVG never get a real
 *      measurement) or a fetch that outright failed.
 *   3. a shared generic placeholder, when neither of the above is known.
 *
 * Both sources are all-or-nothing (never mixes a declared width with an
 * inferred height): `imageNode.width`/`height` are always set as a pair by
 * the HTML parser, and so are an asset's own decoded dimensions.
 */

const DEFAULT_WIDTH_PX = 480; // 5in at 96px/in
const DEFAULT_HEIGHT_PX = 288; // 3in at 96px/in

/** `assetInfo` is whatever an exporter's own per-asset lookup carries — only its `width`/`height` fields are read here. */
export function resolveImagePxSize(imageNode, assetInfo) {
  if (assetInfo?.width && assetInfo?.height) return { widthPx: assetInfo.width, heightPx: assetInfo.height };
  if (imageNode.width && imageNode.height) return { widthPx: imageNode.width, heightPx: imageNode.height };
  return { widthPx: DEFAULT_WIDTH_PX, heightPx: DEFAULT_HEIGHT_PX };
}
