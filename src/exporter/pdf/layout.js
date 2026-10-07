/**
 * PDF page geometry — the points-native (72pt/inch) sibling of DOCX's own
 * `resolvePageGeometry` in `docx/document.js`. Deliberately reimplemented
 * rather than shared: that function isn't exported, and twips (1440/in) vs.
 * points (72/in) are different unit systems end-to-end for what's really
 * only a ~15-line function — duplicating the same *shape* of options
 * handling is more honest than forcing a cross-unit abstraction. The public
 * contract (`pageSize`/`orientation`/`margins` in mm) is identical to
 * DOCX's, so callers get one consistent options shape across every exporter
 * that has physical pages.
 */

// Standard PDF page sizes in points (72pt/inch) — the same format table
// html2pdf.js exposes (it defers to jsPDF's `getPageSize` for this), so a
// `pageSize` value that works there works here too. Values are jsPDF's own
// pt dimensions rounded to the nearest point, matching this file's existing
// integer-pt convention (e.g. A4 595x842, not the unrounded 595.28x841.89).
export const PAGE_SIZES_PT = {
  A0: { w: 2384, h: 3370 },
  A1: { w: 1684, h: 2384 },
  A2: { w: 1191, h: 1684 },
  A3: { w: 842, h: 1191 },
  A4: { w: 595, h: 842 },
  A5: { w: 420, h: 595 },
  A6: { w: 298, h: 420 },
  A7: { w: 210, h: 298 },
  A8: { w: 147, h: 210 },
  A9: { w: 105, h: 147 },
  A10: { w: 74, h: 105 },
  B0: { w: 2835, h: 4008 },
  B1: { w: 2004, h: 2835 },
  B2: { w: 1417, h: 2004 },
  B3: { w: 1001, h: 1417 },
  B4: { w: 709, h: 1001 },
  B5: { w: 499, h: 709 },
  B6: { w: 354, h: 499 },
  B7: { w: 249, h: 354 },
  B8: { w: 176, h: 249 },
  B9: { w: 125, h: 176 },
  B10: { w: 88, h: 125 },
  C0: { w: 2599, h: 3677 },
  C1: { w: 1837, h: 2599 },
  C2: { w: 1298, h: 1837 },
  C3: { w: 918, h: 1298 },
  C4: { w: 649, h: 918 },
  C5: { w: 459, h: 649 },
  C6: { w: 323, h: 459 },
  C7: { w: 230, h: 323 },
  C8: { w: 162, h: 230 },
  C9: { w: 113, h: 162 },
  C10: { w: 79, h: 113 },
  DL: { w: 312, h: 624 },
  Letter: { w: 612, h: 792 },
  GovernmentLetter: { w: 576, h: 756 },
  Legal: { w: 612, h: 1008 },
  JuniorLegal: { w: 576, h: 360 },
  Ledger: { w: 1224, h: 792 },
  Tabloid: { w: 792, h: 1224 },
  CreditCard: { w: 153, h: 243 },
};

export const DEFAULT_MARGIN_PT = 72; // 1in

export function mmToPt(mm) {
  return (mm * 72) / 25.4;
}

/**
 * Resolves `convert()`'s `pageSize`/`orientation`/`margins` options (the
 * same shape DOCX's own page-setup options use) into PDF page dimensions
 * and margins, all in points. `margins.{top,right,bottom,left}` are in mm,
 * each independently defaulting to 1in when omitted — matching DOCX's own
 * per-side default behavior exactly, not just an all-or-nothing default.
 */
export function resolvePageGeometry(options = {}, contentWidthPx, pageMarginPt) {
  const size = PAGE_SIZES_PT[options.pageSize] ?? PAGE_SIZES_PT.A4;
  const landscape = options.orientation === 'landscape';
  let pageWidth = landscape ? size.h : size.w;
  const pageHeight = landscape ? size.w : size.h;

  // A page-width-governing CSS `max-width` on <body> (dom-to-idm.js's own
  // metadata.contentWidthPx) sizes the PDF page to match the width a
  // browser actually renders the page at, instead of a fixed Letter/A4
  // default with no relationship to it — but only when the caller never
  // explicitly chose a page size or orientation: an explicit choice always
  // wins over an inferred one, the same "explicit beats inferred"
  // precedence this project already uses for margin/padding/font-size.
  const autoWidthActive = contentWidthPx && options.pageSize === undefined && options.orientation === undefined;
  // When auto-sizing, the page's own detected left/right padding (see
  // dom-to-idm.js's metadata.pageMarginPt) becomes this project's own
  // left/right margin too, instead of the usual 1in default — a browser
  // has no separate "print margin" layered on top of `.page`'s own edges,
  // so stacking the fixed default on top of a width already derived from
  // those edges would add whitespace the source page never actually had.
  // Only applies to a side the caller didn't explicitly set; an explicit
  // margin always wins, same precedence as the page size/orientation check
  // above.
  const autoLeftRight = autoWidthActive && pageMarginPt !== undefined;

  const margins = options.margins ?? {};
  const top = margins.top !== undefined ? mmToPt(margins.top) : DEFAULT_MARGIN_PT;
  const right = margins.right !== undefined ? mmToPt(margins.right) : autoLeftRight ? pageMarginPt : DEFAULT_MARGIN_PT;
  const bottom = margins.bottom !== undefined ? mmToPt(margins.bottom) : DEFAULT_MARGIN_PT;
  const left = margins.left !== undefined ? mmToPt(margins.left) : autoLeftRight ? pageMarginPt : DEFAULT_MARGIN_PT;

  if (autoWidthActive) {
    pageWidth = contentWidthPx * 0.75 + left + right;
  }

  return {
    pageWidth,
    pageHeight,
    landscape,
    top,
    right,
    bottom,
    left,
    contentWidth: Math.max(1, pageWidth - left - right),
    contentTop: pageHeight - top,
    contentBottom: bottom,
  };
}

// Heading font sizes (pt) — level 1-6, always bold (see fonts.js). Body text
// defaults to 11pt regular, matching DOCX's own Normal-style default so a
// document reads at a consistent scale regardless of which exporter is used.
export const HEADING_SIZES_PT = { 1: 24, 2: 18, 3: 14, 4: 12, 5: 11, 6: 10 };
export const DEFAULT_BODY_FONT_PT = 11;
export const DEFAULT_LINE_HEIGHT_MULTIPLIER = 1.2;

// Fixed indent step per list nesting level — the same "no real layout
// engine, a flat step per depth" simplification pptx/shapes.js's own
// BULLET_STEP_EMU already accepts.
export const LIST_INDENT_STEP_PT = 18;
export const LIST_MARKER_GAP_PT = 6;

export const DEFAULT_TABLE_BORDER_RGB = [0.6, 0.6, 0.6];
export const DEFAULT_TABLE_CELL_PADDING_PT = 4;

export const BLOCK_SPACING_PT = 8; // vertical gap after a block, before the next one

// Every browser applies its own "user agent stylesheet" to raw HTML tags
// BEFORE any of the page's own CSS — see the HTML spec's own suggested
// rendering (https://html.spec.whatwg.org/multipage/rendering.html), which
// Chrome/Firefox/Safari/Edge all implement near-identically. That's why a
// bare `<h2>`/`<p>`/`<ul>`/`<blockquote>`/`<pre>`/`<figure>` with no CSS
// margin of its own still shows visible spacing in every browser. This
// exporter otherwise only ever reads margin the HTML's own CSS explicitly
// declares or a wrapping element bleeds down (see dom-to-idm.js's
// extractMarginPt/mapElement) — so without this table, a document that
// relies on browser defaults instead of declaring its own margin renders
// noticeably tighter here than it does in any browser. Each heading's own
// multiplier is `em` relative to THAT heading's own font-size (the spec's
// `h1{margin:0.67em 0}` is 0.67em of h1's own ~2em size, not the body's) —
// same "em scales with the element's own size" rule HEADING_SIZES_PT already
// encodes per level. Body-level types share the spec's flat `1em` against
// DEFAULT_BODY_FONT_PT. `table` is deliberately absent: the spec gives it no
// default margin at all, only `border-spacing`.
const HEADING_MARGIN_EM = { 1: 0.67, 2: 0.83, 3: 1, 4: 1.33, 5: 1.67, 6: 2.33 };
const BODY_LEVEL_MARGIN_TYPES = new Set(['paragraph', 'list', 'quote', 'code', 'figure']);

/**
 * The browser-default margin for a node's own type — used ONLY as the very
 * last fallback, after any real CSS margin (explicit or bled from a
 * wrapper) has already been checked; explicit HTML/CSS always wins.
 */
export function uaDefaultMarginPt(nodeType, headingLevel) {
  if (nodeType === 'heading') {
    const em = HEADING_MARGIN_EM[headingLevel] ?? HEADING_MARGIN_EM[6];
    const pt = em * (HEADING_SIZES_PT[headingLevel] ?? HEADING_SIZES_PT[6]);
    return { beforePt: pt, afterPt: pt };
  }
  if (BODY_LEVEL_MARGIN_TYPES.has(nodeType)) {
    return { beforePt: DEFAULT_BODY_FONT_PT, afterPt: DEFAULT_BODY_FONT_PT };
  }
  return { beforePt: 0, afterPt: 0 };
}

// A CSS `border-radius`/box-shadow only ever survives into the IDM as a
// yes/no presence flag (see hasBorderRadius/hasBoxShadow's own docs — the
// actual radius/offset/blur/color values aren't representable in OOXML
// either, so this project never carries them) — these are the fixed,
// PDF-native stand-ins for "rounded" and "has a shadow", the same "presence,
// not precision" approximation DOCX/PPTX's own shared shadowEffectXml/
// cardPresetGeometry already make.
export const ROUNDED_CORNER_RADIUS_PT = 8;
export const SHADOW_OFFSET_PT = 3;
export const SHADOW_RGB = [0.75, 0.75, 0.75];
