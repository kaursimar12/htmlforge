/**
 * Intermediate Document Model (IDM).
 *
 * The HTML parser converts HTML into this common model. Every exporter reads
 * from the IDM. This decoupling means adding a new output format = one new
 * exporter with no HTML parsing changes needed.
 */

/* -------------------------------------------------------------------------- */
/* Builder helpers                                                             */
/* -------------------------------------------------------------------------- */

export const doc = (children, metadata) => ({
  type: 'document',
  metadata,
  children,
});

export const section = (children) => ({ type: 'section', children });

export const heading = (level, children) => ({
  type: 'heading',
  level: Math.min(6, Math.max(1, level)),
  children,
});

export const paragraph = (children) => ({
  type: 'paragraph',
  children,
});

export const text = (value) => ({ type: 'text', value });

export const strong = (children) => ({ type: 'strong', children });

export const emphasis = (children) => ({
  type: 'emphasis',
  children,
});

export const codeSpan = (value) => ({ type: 'codeSpan', value });

export const link = (href, children, title) => ({
  type: 'link',
  href,
  title,
  children,
});

export const list = (ordered, items, start) => ({
  type: 'list',
  ordered,
  start,
  items,
});

export const listItem = (children) => ({
  type: 'listItem',
  children,
});

export const tableCell = (children, options = {}) => ({ type: 'tableCell', children, ...options });

// heightPx (CSS px) is set only when the row/its cells explicitly declared a
// height — e.g. a decorative colored strip <td style="height:4px">. It's the
// row's own intended height, not inferred from content.
export const tableRow = (cells, heightPx) => ({
  type: 'tableRow',
  cells,
  heightPx,
});

// width/height (in CSS px) are carried through only when the source markup
// declared them explicitly (an <img>/<svg> width/height attribute) — never
// inferred from decoding image bytes (no npm deps for that). Exporters that
// care about intrinsic size (DOCX) use them when present and fall back to a
// generic default otherwise.
export const image = (src, alt, assetId, width, height) => ({
  type: 'image',
  src,
  alt,
  assetId,
  width,
  height,
});

export const figure = (img, caption) => ({
  type: 'figure',
  image: img,
  caption,
});

// A raster-free `<svg>` — one that draws real vector primitives (rect,
// circle, ellipse, line, polygon, polyline) rather than just wrapping one
// embedded raster `<image>` (see dom-to-idm.js's mapSvgTag/extractSvgShapes
// for that other, older case, which still maps to a plain `image` node
// instead of this one). All shape coordinates are already normalized into
// the same [0,widthPx]×[0,heightPx] px box `widthPx`/`heightPx` describe —
// any source `viewBox` has already been resolved into that box by the
// parser, so no exporter needs to think about viewBox scaling itself.
//
// Each shape in `shapes` is one of:
//   { kind: 'rect', x, y, width, height, fill?, stroke?, strokeWidthPx? }
//   { kind: 'circle', cx, cy, r, fill?, stroke?, strokeWidthPx? }
//   { kind: 'ellipse', cx, cy, rx, ry, fill?, stroke?, strokeWidthPx? }
//   { kind: 'line', x1, y1, x2, y2, stroke?, strokeWidthPx? }
//   { kind: 'polygon'|'polyline', points: [[x,y], ...], fill?, stroke?, strokeWidthPx? }
// `fill`/`stroke` are 6-digit hex strings (no '#') or undefined ('none'/
// unspecified — SVG's own default fill is black, stroke is none, already
// resolved to that default by the parser, not left for exporters to guess).
export const svgShapes = (shapes, widthPx, heightPx) => ({
  type: 'svgShapes',
  shapes,
  widthPx,
  heightPx,
});

// backgroundColor/color are the <pre>'s own explicit CSS (e.g. a Monokai-
// style `pre{background:#272822;color:#eee}` dark theme) — optional, since
// most code blocks rely on the exporter's own default monospace shading
// instead of declaring their own. padding/margin are likewise the <pre>'s
// own explicit CSS box-model properties (see dom-to-idm.js's
// codeFromStyledElement for why these are read directly here rather than
// through the generic wrapper-bleed mechanism every other box-styled node
// type uses).
export const code = (value, language, backgroundColor, color, padding, margin) => ({
  type: 'code',
  language,
  value,
  backgroundColor,
  color,
  padding,
  margin,
});

export const quote = (children) => ({ type: 'quote', children });

/**
 * Wraps inline content that carried explicit character-level formatting on
 * the source element — either CSS (color, background-color, font-weight,
 * font-style, text-decoration, font-size — see src/html-parser/inline-style.js)
 * or a semantic tag with an established formatting convention (<u>, <s>/
 * <del>, <ins>, <mark>, <sub>/<sup>, <var>/<cite> — see dom-to-idm.js). Both
 * are a direct carry-forward of an explicit declaration, the same as
 * StrongNode carries forward an explicit <strong> tag; neither is inferred.
 * `style` shape: { color?, backgroundColor?, bold?, italic?, underline?,
 * strikethrough?, fontSizePt?, fontFamily?, verticalAlign?: 'subscript'|
 * 'superscript' } — always has at least one key when present.
 */
export const styledSpan = (style, children) => ({ type: 'styledSpan', style, children });
