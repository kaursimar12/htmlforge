/**
 * PPTX shape content — text runs/paragraphs for the title & body
 * placeholders, plus standalone picture and table shapes. There's no real
 * layout engine (same constraint the DOCX/XLSX exporters accept): text
 * blocks all collapse into one body placeholder's `<p:txBody>` as a sequence
 * of `<a:p>` paragraphs, while images/tables — which aren't text flow — get
 * their own absolutely-positioned shapes stacked underneath by the caller
 * (slide.js), each sized to fit within its allotted vertical slice.
 */
import { XmlBuilder } from '../../utils/xml-builder.js';
import { flattenInline, isBlockNode } from '../../idm/flatten-text.js';
import { expandTableGrid } from '../shared/table-grid.js';
import { shapeFillXml, shapeLineXml, shadowEffectXml, cardPresetGeometry } from '../shared/shape-style.js';
import { estimateShapeSize, estimateWrappedLines, estimateTextWidth } from '../shared/font-metrics.js';
import { resolveImagePxSize } from '../shared/image-size.js';
import { resolveFontFamilyName } from '../../fonts/resolver.js';
import { PPTX_REL_TYPES } from './layout.js';

/* ─── Inline runs ────────────────────────────────────────────────────────── */

// A real CSS font-size is parsed with the universal 96dpi px->pt conversion
// (see inline-style.js's parseAbsoluteLengthPt) — accurate for print, but a
// slide is a far bigger physical canvas than a browser viewport, so the same
// point size reads much smaller relative to a 13.33in-wide slide than the
// source pixel value did relative to its (usually much narrower) container.
// This reverses that conversion (x4/3) specifically for PPTX runs, so a CSS
// `font-size:40px` becomes 40pt instead of the print-accurate 30pt — a
// deliberate, presentation-specific scale-up, confirmed with the user,
// not a bug in the underlying px->pt math (which stays exact for DOCX/XLSX).
const CSS_TO_SLIDE_FONT_SCALE = 4 / 3;

function scaleFontSize(style) {
  if (style?.fontSizePt == null) return style;
  return { ...style, fontSizePt: style.fontSizePt * CSS_TO_SLIDE_FONT_SCALE };
}

// A page's own `body{font-size:...}` (dom-to-idm.js's own metadata.fontSizePt
// — a document-wide fallback, carried unscaled as a plain CSS pt value
// through rctx.bodyFontSizePt) is what a plain, unstyled paragraph/list
// item/table cell should actually render at — not this file's own flat 18pt
// defaults (DEFAULT_CARD_FONT_PT/DEFAULT_CELL_FONT_PT, "PowerPoint's own
// default autoshape/table-cell text size"), which are only ever the
// fallback for when the source document declares no base size at all.
// Scaled by the same CSS_TO_SLIDE_FONT_SCALE every other CSS-derived font
// size in this file already goes through. `fallback` is genuinely `undefined`
// at one call site (realTableCellXml, when it wants "no explicit `sz`
// attribute at all" rather than a literal 18 to preserve behaviorally
// omitting it) — every other call site passes DEFAULT_CARD_FONT_PT/
// DEFAULT_CELL_FONT_PT.
function resolveDefaultFontPt(rctx, fallback) {
  return rctx?.bodyFontSizePt != null ? rctx.bodyFontSizePt * CSS_TO_SLIDE_FONT_SCALE : fallback;
}

function renderInlineRuns(nodes, rctx, baseProps = {}) {
  return (nodes ?? []).map((n) => renderRun(n, baseProps, rctx)).join('');
}

function renderRun(node, rProps, rctx) {
  switch (node.type) {
    case 'text':
      return textRun(node.value, rProps);
    case 'strong':
      return renderInlineRuns(node.children, rctx, { ...rProps, bold: true });
    case 'emphasis':
      return renderInlineRuns(node.children, rctx, { ...rProps, italic: true });
    case 'codeSpan':
      return textRun(node.value, { ...rProps, code: true });
    case 'styledSpan':
      return renderInlineRuns(node.children, rctx, { ...rProps, ...scaleFontSize(node.style) });
    case 'link':
      return renderHyperlinkRun(node, rProps, rctx);
    case 'lineBreak':
      return '<a:br/>';
    default:
      return '';
  }
}

// An in-document "#anchor" link has no equivalent slide-navigation concept
// wired up here (out of scope for this pass, same as PPTX not supporting
// bookmarks/notes yet) — rendered as plain unlinked text rather than a
// broken/external relationship.
function renderHyperlinkRun(node, rProps, rctx) {
  if (node.href.startsWith('#')) return renderInlineRuns(node.children, rctx, rProps);
  const rId = rctx.relationships.add(PPTX_REL_TYPES.hyperlink, node.href, 'External');
  return renderInlineRuns(node.children, rctx, { ...rProps, hyperlinkRId: rId });
}

function textRun(value, rProps = {}) {
  const b = new XmlBuilder();
  b.el('a:r');
  b.raw(renderRPr(rProps));
  b.el('a:t');
  b.text(value);
  b.up();
  b.up();
  return b.end();
}

// ST_TextCharacterProperties baseline shift, in thousandths of a percent —
// there's no "true" font-relative value the way CSS vertical-align has; 30%
// up / 25% down are the same fixed proportions PowerPoint's own UI applies
// when you toggle superscript/subscript.
const SUPERSCRIPT_BASELINE = 30000;
const SUBSCRIPT_BASELINE = -25000;

function renderRPr(rProps) {
  const b = new XmlBuilder();
  b.el('a:rPr', {
    lang: 'en-US',
    b: rProps.bold ? 1 : undefined,
    i: rProps.italic ? 1 : undefined,
    u: rProps.underline ? 'sng' : undefined,
    strike: rProps.strikethrough ? 'sngStrike' : undefined,
    sz: rProps.fontSizePt ? Math.round(rProps.fontSizePt * 100) : undefined,
    baseline: rProps.verticalAlign === 'superscript' ? SUPERSCRIPT_BASELINE
      : rProps.verticalAlign === 'subscript' ? SUBSCRIPT_BASELINE : undefined,
  });
  // Child element order matters (CT_TextCharacterProperties is a strict
  // sequence): fill, then highlight, then font, then hlinkClick last.
  if (rProps.color) {
    b.el('a:solidFill');
    b.self('a:srgbClr', { val: rProps.color });
    b.up();
  }
  if (rProps.backgroundColor) {
    b.el('a:highlight');
    b.self('a:srgbClr', { val: rProps.backgroundColor });
    b.up();
  }
  if (rProps.code) {
    b.self('a:latin', { typeface: 'Courier New' });
  } else if (rProps.fontFamily) {
    // A bare CSS generic keyword ("sans-serif") has no font literally named
    // that for PowerPoint to find — mapped to a real installable name first.
    b.self('a:latin', { typeface: resolveFontFamilyName(rProps.fontFamily) });
  }
  if (rProps.hyperlinkRId) b.self('a:hlinkClick', { 'r:id': rProps.hyperlinkRId });
  b.up();
  return b.end();
}

/* ─── Body placeholder paragraphs ────────────────────────────────────────── */

// marL/indent for a bullet at nesting `depth` (0-based) — a fixed step per
// level, the same "no real layout engine, just consistent proportions"
// approach as computeSlideGeometry.
const BULLET_STEP_EMU = 457200; // 0.5in
const BULLET_BASE_MARL = 228600; // 0.25in
const BULLET_HANG_INDENT = -228600;

// node.align/cell.align are already normalized to one of these four CSS
// keywords by the HTML parser (see inline-style.js's ALIGNMENTS) — DrawingML
// just spells them differently on `a:pPr algn`.
const ALIGN_MAP = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };

// A HeadingNode with no explicit CSS font-size otherwise renders as merely
// bold text at the placeholder's own default size — indistinguishable from
// an h3 next to an h6, since our master/layout's `a:lstStyle` is empty (no
// `p:txStyles` level-based sizing the way a real PowerPoint template has).
// This default only ever applies where nothing more specific already does:
// an h1/h2/h3 CSS rule wrapping the heading's own children in a styledSpan
// (see dom-to-idm.js's mapInlineWithStyle) always overrides it, the same
// "explicit CSS beats inference" precedence used throughout this project.
const DEFAULT_HEADING_SIZE_PT = { 1: 32, 2: 28, 3: 24, 4: 20, 5: 18, 6: 16 };

// The estimated box height this paragraph sits in (see slide.js's
// estimateTextRunHeightEmu / this file's estimateCardBlockHeightPt) assumes
// an exact `fontSizePt * 1.2` line height — but that's only ever OUR
// assumption; left unset, the actual line spacing PowerPoint/LibreOffice/
// Google Slides applies at open-time is up to that app's own font metrics,
// which can run measurably taller than 1.2x, especially once a line mixes in
// a script (CJK/Devanagari/Arabic/emoji) that pulls in a fallback font with
// its own taller natural line-height. Pinning `a:lnSpc` to an absolute point
// value (not a percentage — a percentage is still relative to the font's own
// natural line height, the exact variance this exists to remove) forces the
// real render to match the estimate exactly, the same "make the render
// match the assumption" fix already applied to font size (see DEFAULT_
// CARD_FONT_PT / DEFAULT_HEADING_SIZE_PT's own explicit `sz` everywhere).
function lnSpcXml(fontSizePt, multiplier = 1.2) {
  return `<a:lnSpc><a:spcPts val="${Math.round(fontSizePt * multiplier * 100)}"/></a:lnSpc>`;
}

function paragraph(runsXml, { pPrAttrs, bullet, lnSpcPt } = {}) {
  const b = new XmlBuilder();
  b.el('a:p');
  // marL/indent default to 0 unless a caller explicitly overrides them (e.g.
  // renderListParagraphs's own hanging-indent marL/BULLET_HANG_INDENT for a
  // real bullet). Our own slide master/layout never define a `p:txStyles`
  // (see layout.js), so a paragraph that doesn't set these itself inherits
  // PowerPoint's own BUILT-IN default level-1 list style — a hanging indent
  // meant for bullets (marL positive, indent negative so a bullet's first
  // line nets out flush). OOXML's hanging-indent rule applies `marL+indent`
  // only to the paragraph's very FIRST line; every line after that —
  // including a hard `<a:br/>` break within the SAME paragraph, not just a
  // wrapped continuation — uses `marL` alone. A plain multi-line paragraph
  // with no bullet (e.g. `<b>Invoice No.</b> : ...<br><b>Date</b> : ...`)
  // was inheriting that ambient default with nothing here to override it:
  // its first line rendered flush (marL+indent ≈ 0) while every line
  // after a `<br>` rendered visibly indented (marL alone) — a real,
  // observed PowerPoint rendering difference, not a stray leading space.
  b.el('a:pPr', { marL: 0, indent: 0, ...pPrAttrs });
  // CT_TextParagraphProperties is a strict child sequence — lnSpc must come
  // before the bu*/buNone element, same ordering constraint renderRPr's own
  // comment on CT_TextCharacterProperties already documents.
  if (lnSpcPt) b.raw(lnSpcXml(lnSpcPt));
  if (bullet) b.raw(bullet);
  else b.self('a:buNone');
  b.up();
  b.raw(runsXml || textRun('', {}));
  b.up();
  return b.end();
}

function renderListParagraphs(list, depth, rctx) {
  const out = [];
  const marL = BULLET_BASE_MARL + depth * BULLET_STEP_EMU;
  const fontPt = resolveDefaultFontPt(rctx, DEFAULT_CARD_FONT_PT);
  list.items.forEach((item, i) => {
    const [first, ...rest] = item.children ?? [];
    const runsXml = first?.type === 'paragraph' ? renderInlineRuns(first.children, rctx, { fontSizePt: fontPt }) : '';
    const bullet = list.ordered
      ? `<a:buAutoNum type="arabicPeriod"${i === 0 && list.start && list.start !== 1 ? ` startAt="${list.start}"` : ''}/>`
      : '<a:buChar char="•"/>';
    out.push(paragraph(runsXml, { pPrAttrs: { marL, indent: BULLET_HANG_INDENT, lvl: depth || undefined }, bullet, lnSpcPt: fontPt }));
    if (first && first.type !== 'paragraph') out.push(...renderBodyBlockParagraphs(first, rctx));
    for (const child of rest) {
      if (child.type === 'list') out.push(...renderListParagraphs(child, depth + 1, rctx));
      else out.push(...renderBodyBlockParagraphs(child, rctx));
    }
  });
  return out;
}

// Renders one text-flow block to one or more `<a:p>` paragraphs for the body
// placeholder. Images/figures/tables are NOT handled here — slide.js pulls
// those out beforehand and renders them as their own separate shapes.
export function renderBodyBlockParagraphs(node, rctx) {
  const algn = ALIGN_MAP[node.align];
  // A <div class="..."> with its own background (e.g. `.circle{background:
  // #2563EB;color:white}`) stamps `backgroundColor` directly on the
  // paragraph/heading it wraps — a block-level property, entirely separate
  // from the run-level styledSpan a same-element `color` produces (see
  // dom-to-idm.js's mapElement). DrawingML text bodies have no per-paragraph
  // fill the way Word has `w:shd`, so without this, that block's own text
  // color (often white, chosen specifically to sit on that background) has
  // nothing to sit on and silently disappears against the slide's own
  // background. Applying it as a run highlight is an approximation — a tight
  // swatch behind the text, not a padded/rounded box — but it's what keeps
  // the text from going invisible, using the same mechanism already used for
  // <mark> and the code block's own background.
  const bg = node.backgroundColor ? { backgroundColor: node.backgroundColor } : undefined;
  const fontPt = resolveDefaultFontPt(rctx, DEFAULT_CARD_FONT_PT);
  switch (node.type) {
    case 'paragraph':
      return [paragraph(renderInlineRuns(node.children, rctx, { fontSizePt: fontPt, ...bg }), { pPrAttrs: { algn }, lnSpcPt: fontPt })];
    case 'heading': {
      const fontSizePt = DEFAULT_HEADING_SIZE_PT[node.level];
      return [paragraph(renderInlineRuns(node.children, rctx, { bold: true, fontSizePt, ...bg }), { pPrAttrs: { algn }, lnSpcPt: fontSizePt })];
    }
    // A blockquote's own `children` are block nodes (typically one paragraph
    // — see normalizer's normalizeBlock), not inline content, so they can't
    // go through renderInlineRuns directly the way a plain paragraph's
    // `children` can — doing so silently produced empty `<a:t/>` runs (every
    // paragraph/heading/etc. node type falls through renderRun's `default:
    // return ''`), losing the quote's text entirely. Each paragraph child
    // renders with the standard italic + indent treatment (plus the quote's
    // own backgroundColor, e.g. `blockquote{background:#eef4ff}`, the same
    // highlight-run approximation the 'paragraph'/'heading' cases above use);
    // any other child (a nested list/table/quote) recurses as itself, same
    // as docx/document.js's renderQuoteChild.
    case 'quote':
      return node.children.flatMap((child) => {
        if (child.type !== 'paragraph') return renderBodyBlockParagraphs(child, rctx);
        const childAlgn = ALIGN_MAP[child.align] ?? algn;
        return [paragraph(renderInlineRuns(child.children, rctx, { italic: true, fontSizePt: fontPt, ...bg }), { pPrAttrs: { marL: BULLET_BASE_MARL, algn: childAlgn }, lnSpcPt: fontPt })];
      });
    case 'code':
      return node.value.split('\n').map((line) => paragraph(textRun(line, { code: true, fontSizePt: fontPt, color: node.color, backgroundColor: node.backgroundColor }), { lnSpcPt: fontPt }));
    case 'list':
      return renderListParagraphs(node, 0, rctx);
    default:
      return [];
  }
}

/* ─── Images ─────────────────────────────────────────────────────────────── */

const EMU_PER_PX = 9525; // matches docx/document.js

function imageNaturalSize(imageNode, assetInfo) {
  const { widthPx, heightPx } = resolveImagePxSize(imageNode, assetInfo);
  return { cx: widthPx * EMU_PER_PX, cy: heightPx * EMU_PER_PX };
}

function fitWithin(size, maxCx, maxCy) {
  const scale = Math.min(1, maxCx / size.cx, maxCy / size.cy);
  return { cx: Math.round(size.cx * scale), cy: Math.round(size.cy * scale) };
}

// Every slide has its own slideN.xml.rels — an image relationship id from
// one slide means nothing on another, so (unlike DOCX's single shared
// RelationshipManager) each slide must register its own rId for an asset the
// first time that slide's content actually uses it. `rctx.images` caches
// that per-slide registration so the same asset reused twice in one slide
// doesn't get two rIds.
function resolveImageRel(imageNode, rctx) {
  if (!imageNode.assetId) return undefined;
  if (rctx.images.has(imageNode.assetId)) return rctx.images.get(imageNode.assetId);
  const info = rctx.assetInfoByAssetId.get(imageNode.assetId);
  if (!info) return undefined;
  const relId = rctx.relationships.add(PPTX_REL_TYPES.image, `../media/${info.filename}`);
  const entry = { relId, width: info.width, height: info.height };
  rctx.images.set(imageNode.assetId, entry);
  return entry;
}

// Returns `{ xml, cy }` (cy so the caller can advance its y-cursor past this
// shape) — an unresolved image (fetch failed / no assetId) renders nothing,
// same as DOCX: resolveAssets already emitted its own ctx.warn for the
// failure, so this doesn't warn a second time.
// `verticalAlign` defaults to 'top' — correct for a standalone flow image
// (`box.cy` there is whatever vertical space remains on the slide/in the
// current layout slot, not a fixed container the image should center
// inside; centering within LEFTOVER page space would drift the image down
// from where the preceding content actually ended). A table cell is
// different: `box` there is the cell's own real, fixed bounds (see
// buildTableGraphicFrame), so 'middle' centers the image within it, same
// as `text-align:center` already centers it horizontally either way.
export function renderPictureShape(imageNode, rctx, box, shapeId, verticalAlign = 'top') {
  const asset = resolveImageRel(imageNode, rctx);
  if (!asset) return { xml: '', cy: 0 };
  const { cx, cy } = fitWithin(imageNaturalSize(imageNode, asset), box.cx, box.cy);
  const x = box.x + Math.round((box.cx - cx) / 2);
  const y = verticalAlign === 'middle' ? box.y + Math.round((box.cy - cy) / 2) : box.y;

  const b = new XmlBuilder();
  b.el('p:pic');
  b.el('p:nvPicPr');
  b.self('p:cNvPr', { id: shapeId, name: imageNode.alt || `Picture ${shapeId}`, descr: imageNode.alt || undefined });
  b.self('p:cNvPicPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:blipFill');
  b.self('a:blip', { 'r:embed': asset.relId });
  b.el('a:stretch');
  b.self('a:fillRect');
  b.up();
  b.up();
  b.el('p:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x, y });
  b.self('a:ext', { cx, cy });
  b.up();
  b.el('a:prstGeom', { prst: 'rect' });
  b.self('a:avLst');
  b.up();
  b.up();
  b.up();
  return { xml: b.end(), cy };
}

/* ─── SVG vector shapes ──────────────────────────────────────────────────── */
// One SVG (`svgShapes`, see dom-to-idm.js's extractSvgShapes) becomes one
// separately-positioned `<p:sp>` per shape — unlike DOCX (flowed paragraph
// content with no absolute positioning), a PPTX slide's own shape tree
// (`p:spTree`) already places every shape via explicit `a:off`/`a:ext` EMU
// coordinates, so no special anchoring mechanism is needed to composite
// several shapes at their correct relative positions the way DOCX does —
// this is simply the same absolute placement every other shape here
// already uses. PPTX/DrawingML's own coordinate space is Y-down from the
// top-left, the exact same convention SVG itself uses, so — unlike PDF's
// Y-up page space — no vertical flip is needed, only scaling and origin
// translation.


function prstGeomSvgShapeXml(prst, x, y, cx, cy, shape, shapeId, flip) {
  const b = new XmlBuilder();
  b.el('p:sp');
  b.el('p:nvSpPr');
  b.self('p:cNvPr', { id: shapeId, name: `Shape ${shapeId}` });
  b.self('p:cNvSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:spPr');
  b.el('a:xfrm', { flipH: flip?.h ? 1 : undefined, flipV: flip?.v ? 1 : undefined });
  b.self('a:off', { x: Math.round(x), y: Math.round(y) });
  b.self('a:ext', { cx: Math.max(1, Math.round(cx)), cy: Math.max(1, Math.round(cy)) });
  b.up();
  b.el('a:prstGeom', { prst });
  b.self('a:avLst');
  b.up();
  b.raw(shapeFillXml({ backgroundColor: shape.fill }));
  if (shape.stroke) b.raw(shapeLineXml({ color: shape.stroke, sizeEighthPt: Math.max(2, Math.round((shape.strokeWidthPx ?? 1) * 8)) }));
  b.up(); // p:spPr
  b.el('p:txBody');
  b.self('a:bodyPr');
  b.self('a:lstStyle');
  b.up();
  b.up();
  return b.end();
}

// A polygon/polyline has no DrawingML preset geometry — built as a real
// `a:custGeom` custom path instead, in the path's own local coordinate
// space (its `w`/`h` attributes define that space; DrawingML scales it to
// the shape's actual `a:ext` on the slide automatically, so the path's own
// units don't need to match EMU). `a:close` is only emitted for a polygon
// (SVG auto-closes it); a polyline's own last point stays open, matching
// its real stroke behavior (see pdf/page-flow.js's drawPolygon for the
// identical fill-still-implicitly-closes/stroke-stays-open distinction —
// DrawingML's own path painting follows the same convention).
function customGeomSvgShapeXml(points, closed, x, y, cx, cy, shape, shapeId) {
  const localPoints = points.map(([px, py]) => [Math.round(px), Math.round(py)]);
  const w = Math.max(1, Math.round(cx));
  const h = Math.max(1, Math.round(cy));

  const b = new XmlBuilder();
  b.el('p:sp');
  b.el('p:nvSpPr');
  b.self('p:cNvPr', { id: shapeId, name: `Shape ${shapeId}` });
  b.self('p:cNvSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x: Math.round(x), y: Math.round(y) });
  b.self('a:ext', { cx: w, cy: h });
  b.up();
  b.el('a:custGeom');
  b.el('a:pathLst');
  b.el('a:path', { w, h });
  localPoints.forEach(([px, py], i) => {
    b.el(i === 0 ? 'a:moveTo' : 'a:lnTo');
    b.self('a:pt', { x: px, y: py });
    b.up();
  });
  if (closed) b.self('a:close');
  b.up(); // a:path
  b.up(); // a:pathLst
  b.up(); // a:custGeom
  b.raw(shapeFillXml({ backgroundColor: shape.fill }));
  if (shape.stroke) b.raw(shapeLineXml({ color: shape.stroke, sizeEighthPt: Math.max(2, Math.round((shape.strokeWidthPx ?? 1) * 8)) }));
  b.up(); // p:spPr
  b.el('p:txBody');
  b.self('a:bodyPr');
  b.self('a:lstStyle');
  b.up();
  b.up();
  return b.end();
}

function svgShapeXml(shape, originX, originY, scale, shapeId) {
  switch (shape.kind) {
    case 'rect':
      return prstGeomSvgShapeXml('rect', originX + shape.x * scale, originY + shape.y * scale, shape.width * scale, shape.height * scale, shape, shapeId);
    case 'circle':
      return prstGeomSvgShapeXml('ellipse', originX + (shape.cx - shape.r) * scale, originY + (shape.cy - shape.r) * scale, shape.r * 2 * scale, shape.r * 2 * scale, shape, shapeId);
    case 'ellipse':
      return prstGeomSvgShapeXml('ellipse', originX + (shape.cx - shape.rx) * scale, originY + (shape.cy - shape.ry) * scale, shape.rx * 2 * scale, shape.ry * 2 * scale, shape, shapeId);
    case 'line': {
      const x = originX + Math.min(shape.x1, shape.x2) * scale;
      const y = originY + Math.min(shape.y1, shape.y2) * scale;
      const cx = Math.abs(shape.x2 - shape.x1) * scale;
      const cy = Math.abs(shape.y2 - shape.y1) * scale;
      // DrawingML's `line` preset geometry always draws its diagonal from
      // top-left to bottom-right of its own bounding box — flipH/flipV
      // mirror it to instead run top-right-to-bottom-left (or reversed)
      // when the real endpoints go the other way.
      const flip = { h: (shape.x2 < shape.x1) !== (shape.y2 < shape.y1), v: false };
      return prstGeomSvgShapeXml('line', x, y, cx, cy, shape, shapeId, flip);
    }
    case 'polygon':
    case 'polyline': {
      const scaled = shape.points.map(([x, y]) => [originX + x * scale, originY + y * scale]);
      const minX = Math.min(...scaled.map((p) => p[0]));
      const minY = Math.min(...scaled.map((p) => p[1]));
      const maxX = Math.max(...scaled.map((p) => p[0]));
      const maxY = Math.max(...scaled.map((p) => p[1]));
      const localPoints = scaled.map(([x, y]) => [x - minX, y - minY]);
      return customGeomSvgShapeXml(localPoints, shape.kind === 'polygon', minX, minY, maxX - minX, maxY - minY, shape, shapeId);
    }
    default:
      return '';
  }
}

// Returns `{ xml, cy }` for every shape in one SVG, aspect-fit-scaled and
// centered within `box` — same "reserve a slice, shrink to fit, center
// horizontally" pattern renderPictureShape uses for raster images.
// `nextShapeId` is a callback (not a plain number) because this draws a
// variable number of shapes, one id each — matching renderLayoutBlock's own
// `() => shapeId++` convention for the same reason.
export function renderSvgShapesShape(node, rctx, box, nextShapeId) {
  const { cx, cy } = fitWithin({ cx: node.widthPx * EMU_PER_PX, cy: node.heightPx * EMU_PER_PX }, box.cx, box.cy);
  const originX = box.x + Math.round((box.cx - cx) / 2);
  const originY = box.y;
  const scale = cx / node.widthPx; // EMU per SVG-unit — uniform, aspect always preserved

  const xml = node.shapes.map((shape) => svgShapeXml(shape, originX, originY, scale, nextShapeId())).join('');
  return { xml, cy };
}

/* ─── Styled boxes (badges/pills) ────────────────────────────────────────── */
// A short inline element with its own background — `<p class="badge">
// VERSION 1.0</p>`, a status pill, a tag — reads as a real rounded, padded
// pill in a browser (`display:inline-block` shrink-wraps to its own
// content), not a paragraph that happens to span the slide's full width.
// The plain run-highlight approximation used for a *long* background block
// (renderBodyBlockParagraphs's own `bg` handling — a gradient panel, a
// colored card) reads fine as prose with a tinted backdrop, but wrapping a
// two-word badge in it just puts a tight rectangle behind the glyphs with no
// padding or rounding — nothing like the pill it's meant to be. This
// promotes it to a real shape instead, sized to its own text rather than
// stretched across the content column, which is exactly what the "is this
// short enough to be a badge, not a paragraph" heuristic below exists to
// detect (there's no `display:inline-block` signal anywhere in the IDM to
// key off directly).

// Only a short, single-line-ish block is treated as a badge/pill — a long
// background block (a gradient panel's full sentence, a card) still reads
// correctly as prose with a tinted backdrop via renderBodyBlockParagraphs's
// own run-highlight handling, and isn't promoted to its own shrink-wrapped
// shape (which would look wrong stretched, or absurd if forced to shrink-wrap
// a whole paragraph down to one line).
const BADGE_MAX_CHARS = 40;

export function isStyledBoxNode(node) {
  // A plain solid backgroundColor is NOT enough on its own: unlike border/
  // backgroundGradient/shrinkToFit (all rare, deliberate signals), a solid
  // background is exactly what bleeds down from a big, unrelated page-level
  // wrapper's own backdrop (`.slide{background:white}`) onto EVERY plain
  // heading/paragraph inside it that has no styling of its own — see
  // dom-to-idm.js's mapElement. Without this second check, every short
  // section heading in such a wrapper (which is short text — headings
  // usually are) got wrongly promoted into its own full-width "card" shape,
  // picking up the wrapper's shadow along with it. Requiring one of these
  // stronger signals keeps the true positives (a real bordered box, a real
  // gradient panel, a real inline-block/fixed-width badge) while dropping
  // "just happens to sit inside something with a background." `shadow` is
  // included here too now — unlike backgroundColor, dom-to-idm.js's own
  // shadowRoundedTrusted gate already only lets `shadow` reach a node when
  // it's either the wrapper's sole child or the wrapper also has a real
  // border, so a `.slide`-style many-children/no-border wrapper's shadow
  // never reaches here at all, while a real single-child
  // `.shadow{background:white;box-shadow:...}` div's does.
  //
  // `hasRealBorder` covers a per-side border (border-left, the common
  // "callout/admonition" accent-stripe pattern — a documentation site's
  // Warning/Note/Tip box) as well as a uniform one — previously only
  // `node.border` (uniform) counted at all here, so a border-left-only
  // callout didn't even clear the first two conditions, on top of the length
  // cap below. The BADGE_MAX_CHARS cap itself is skipped whenever there's a
  // real border: that length limit exists only to keep a *plain background*
  // from promoting an unrelated long paragraph (the false-positive this
  // whole function guards against) — a real border is already the strong,
  // deliberate "this is one bounded panel" signal the rest of this project
  // trusts regardless of length (see cardGroupId, DOCX's isOwnBox), so a
  // long callout box shouldn't be held to a limit meant for short badges.
  // Skipping it is also the only way such a callout ever gets a real box at
  // all: DrawingML text bodies have no per-paragraph border/shading the way
  // Word's w:pBdr/w:shd do, so without promotion to its own shape, a long
  // bordered paragraph falls back to a tight run-highlight with no border
  // whatsoever, wrapped mid-shape — visibly broken, not just imprecise.
  const hasRealBorder = !!node.border || !!node.borderTop || !!node.borderRight ||
    !!node.borderBottom || !!node.borderLeft;
  // `node.padding` (own explicit CSS padding, e.g. `<p style="background:
  // #1f6586;padding:20px">`) is included alongside border/gradient/
  // shrinkToFit/shadow as a second, equally-trustworthy signal: a plain
  // `<div>`'s own real padding only bleeds onto a child at all when that
  // child is either the wrapper's sole child, or the wrapper is itself a
  // trusted panel (see dom-to-idm.js's shadowRoundedTrusted) — a big
  // many-children page wrapper's own padding never reaches an unrelated
  // grandchild the promiscuous way backgroundColor does — so "this node has
  // its own padding AND a background" is a reliable "deliberate call-out
  // banner" signal, not a false positive from an unrelated ancestor's
  // backdrop. Needed because DrawingML's text body has no per-paragraph
  // shading the way Word's w:shd does: a plain background-only paragraph
  // with no other signal falls back to a tight run-highlight (see
  // renderBodyBlockParagraphs' own `bg` handling) that can never look like
  // a full-width banner, no matter which app opens the file — unlike DOCX's
  // w:shd, DrawingML's `a:highlight` is inherently glyph-hugging.
  return (node.type === 'paragraph' || node.type === 'heading') &&
    (!!node.backgroundColor || hasRealBorder || !!node.backgroundGradient) &&
    (hasRealBorder || !!node.backgroundGradient || !!node.shrinkToFit || !!node.shadow || !!node.padding) &&
    (hasRealBorder || flattenInline(node.children ?? []).length <= BADGE_MAX_CHARS);
}

const EMU_PER_PT = 12700;

/* ─── Thematic breaks (<hr>) ─────────────────────────────────────────────── */
// A bare `<hr>` (ThematicBreakNode — see dom-to-idm.js's own mapping, which
// deliberately carries no color/width of its own, matching PDF's own
// renderThematicBreak: a flat generic gray line regardless of any custom
// `border-top` the source `<hr>` declared) used to fall through
// classifyContentUnits' own final `pushText` branch — MEDIA_TYPES had no
// entry for it, so it silently became (or joined) a plain "text" unit, and
// renderBodyBlockParagraphs' own switch has no 'thematicBreak' case either
// (falls to its `default: return []`). A standalone `<hr>` sitting between
// two `<table>`s — this project's own invoice, between the header and
// "From"/"Bill To" tables — thus became its own EMPTY text unit that still
// claimed the slide's one real body placeholder (the first text unit on the
// slide always does — see renderSlidePage's own note), rendering a
// placeholder shape with ZERO actual `<a:p>` content: PowerPoint shows an
// unfilled placeholder's own "Click to add text" prompt for that, visibly
// overlapping whatever else sits nearby. Rendering it as a real thin-line
// shape here (a MEDIA-like unit, see MEDIA_TYPES/renderMediaShape) fixes
// both problems at once — actual visible content instead of nothing, and
// never competing for the body placeholder the way a text unit does.
const THEMATIC_BREAK_HEIGHT_EMU = Math.round(EMU_PER_PT * 6); // a small gap above/below the line itself, same category as PDF's own `gap` constant
const THEMATIC_BREAK_LINE_RGB = '999999';

export function renderThematicBreakShape(box, shapeId) {
  const y = box.y + Math.round(THEMATIC_BREAK_HEIGHT_EMU / 2);
  const b = new XmlBuilder();
  b.el('p:sp');
  b.el('p:nvSpPr');
  b.self('p:cNvPr', { id: shapeId, name: `Divider ${shapeId}` });
  b.self('p:cNvSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x: box.x, y });
  b.self('a:ext', { cx: box.cx, cy: 9525 }); // 1px-equivalent — a real, if hairline, height (0 would be an invalid/degenerate extent)
  b.up();
  b.el('a:prstGeom', { prst: 'rect' });
  b.self('a:avLst');
  b.up();
  b.el('a:solidFill');
  b.self('a:srgbClr', { val: THEMATIC_BREAK_LINE_RGB });
  b.up();
  b.up(); // p:spPr
  b.el('p:txBody');
  b.self('a:bodyPr');
  b.self('a:lstStyle');
  b.up();
  b.up();
  return { xml: b.end(), cy: THEMATIC_BREAK_HEIGHT_EMU };
}

const DEFAULT_BOX_PADDING_PT = { top: 6, right: 10, bottom: 6, left: 10 };
const DEFAULT_BOX_FONT_PT = 14;

// A whole-element inline style (`mapInlineWithStyle` wrapping every child in
// one styledSpan — the shape produced by `<p class="badge">`, matching
// DOCX's own cellFontScale heuristic) is the only place a badge's real font
// size lives; anything else falls back to a plain, readable default.
function resolveOwnFontSizePt(node, fallback = DEFAULT_BOX_FONT_PT) {
  const only = node.children?.length === 1 ? node.children[0] : undefined;
  if (only?.type === 'styledSpan' && only.style?.fontSizePt) {
    return only.style.fontSizePt * CSS_TO_SLIDE_FONT_SCALE;
  }
  return fallback;
}

// Same lookup as resolveOwnFontSizePt, for the font-family that
// font-metrics.js needs to pick the right per-character width table. Falls
// back to the deck's own default typeface (see layout.js's themeXml) — the
// same one plain unstyled text actually renders in.
function resolveOwnFontFamily(node, fallback = 'Calibri') {
  const only = node.children?.length === 1 ? node.children[0] : undefined;
  if (only?.type === 'styledSpan' && only.style?.fontFamily) return resolveFontFamilyName(only.style.fontFamily);
  return fallback;
}

// Returns `{ xml, cy }` sized to fit `node`'s own (short) text rather than
// stretched to `box.cx` — `box` here is only the outer bound this shape may
// not exceed, not its actual size.
export function renderStyledBoxShape(node, rctx, box, shapeId) {
  const text = flattenInline(node.children ?? []);
  const fontSizePt = resolveOwnFontSizePt(node);
  const fontFamily = resolveOwnFontFamily(node);
  const paddingPt = {
    top: node.padding?.top ?? DEFAULT_BOX_PADDING_PT.top,
    right: node.padding?.right ?? DEFAULT_BOX_PADDING_PT.right,
    bottom: node.padding?.bottom ?? DEFAULT_BOX_PADDING_PT.bottom,
    left: node.padding?.left ?? DEFAULT_BOX_PADDING_PT.left,
  };
  const padLeft = paddingPt.left * EMU_PER_PT;
  const padRight = paddingPt.right * EMU_PER_PT;
  const padTop = paddingPt.top * EMU_PER_PT;
  const padBottom = paddingPt.bottom * EMU_PER_PT;

  // A badge has no fixed width to wrap to — font-metrics.js sizes it to one
  // line's own estimated width (with a built-in bias margin, since the
  // per-character tables are an estimate, not a measurement).
  const estimated = estimateShapeSize(text, fontSizePt, fontFamily, { padding: paddingPt });

  // An explicit CSS width/height (e.g. `.circle{width:120px;height:120px}`)
  // is a real, authored size — trust it over the estimate above, the same
  // "explicit beats inferred" precedence used throughout this project.
  const cx = Math.min(box.cx, node.widthPx != null ? Math.round(node.widthPx * EMU_PER_PX) : Math.round(estimated.width * EMU_PER_PT));
  const cy = Math.min(box.cy, node.heightPx != null ? Math.round(node.heightPx * EMU_PER_PX) : Math.round(estimated.height * EMU_PER_PT));

  const b = new XmlBuilder();
  b.el('p:sp');
  b.el('p:nvSpPr');
  b.self('p:cNvPr', { id: shapeId, name: `Box ${shapeId}` });
  b.self('p:cNvSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x: box.x, y: box.y });
  b.self('a:ext', { cx, cy });
  b.up();
  // A badge defaults to a soft pill look (roundRect) even with no explicit
  // border-radius (see isStyledBoxNode's own "reads as a real pill" note) —
  // but a genuine CSS circle/oval (border-radius:50%+, e.g. a colored status
  // dot) should render as a real ellipse instead, not a rounded rectangle.
  b.el('a:prstGeom', { prst: node.circular ? 'ellipse' : 'roundRect' });
  b.self('a:avLst');
  b.up();
  b.raw(shapeFillXml(node));
  b.raw(shapeLineXml(node.border));
  if (node.shadow) b.raw(shadowEffectXml());
  b.up(); // p:spPr
  b.el('p:txBody');
  b.el('a:bodyPr', {
    wrap: 'none', anchor: 'ctr',
    lIns: Math.round(padLeft), tIns: Math.round(padTop), rIns: Math.round(padRight), bIns: Math.round(padBottom),
  });
  b.self('a:spAutoFit');
  b.up();
  b.self('a:lstStyle');
  // backgroundColor is already the shape's own fill above — rendering the
  // paragraph with it still set would additionally highlight the run itself
  // (renderBodyBlockParagraphs's own `bg` handling), redundant now that a
  // real shape sits behind it.
  b.raw(renderBodyBlockParagraphs({ ...node, backgroundColor: undefined }, rctx).join(''));
  b.up();
  b.up(); // p:sp
  return { xml: b.end(), cy };
}

/* ─── Cards (bordered/filled multi-block containers) ────────────────────── */
// `<div class="card">` wrapping a heading plus a couple of paragraphs — see
// dom-to-idm.js's mapElement, which stamps every node this kind of container
// produces with a shared `cardGroup` token. Unlike a badge (one short line,
// shrink-wrapped to its own text), a card keeps the content column's full
// width and grows to fit however many blocks it contains, since HTML's own
// `<div>` is block-level by default.

const DEFAULT_CARD_FONT_PT = 18; // PowerPoint's own default autoshape text size
const DEFAULT_CARD_PADDING_PT = { top: 12, right: 16, bottom: 12, left: 16 };
// Exported alongside estimateCardBlockHeightPt so pptx/slide.js's own
// body-placeholder text-run height estimate uses the exact same per-block
// gap, not a second constant that could silently drift from this one.
export const CARD_BLOCK_SPACING_PT = 6;

// `maxWidthPt`, when known, lets a paragraph/heading/quote wrap across
// several lines via font-metrics.js instead of the flat one-line guess —
// real prose in a card is exactly the case that most needs it (a two-
// sentence card blurb is very often 2-3 lines, not 1). Lists/code still use
// one line per item/source-line — wrapping *within* a bullet or a code line
// isn't modeled here, the same "not a full layout engine" scope limit the
// rest of this exporter accepts. `maxWidthPt` is omitted by callers that
// don't yet know the box width (kept for backward compatibility) — those
// keep the old one-line-per-block behavior.
//
// Exported so pptx/slide.js's own body-placeholder text runs (plain
// paragraphs/headings/lists interleaved with images/cards in the body,
// not inside a card) can estimate their own height too, using the exact
// same per-block-type math and DEFAULT_CARD_FONT_PT — PowerPoint's own
// default autoshape/placeholder text size either way — rather than a
// second, separately-drifting estimator.
export function estimateCardBlockHeightPt(node, maxWidthPt, rctx) {
  const fontPt = resolveDefaultFontPt(rctx, DEFAULT_CARD_FONT_PT);
  if (node.type === 'heading') {
    const fontSizePt = DEFAULT_HEADING_SIZE_PT[node.level] ?? fontPt;
    return estimateBlockLineCount(node, maxWidthPt, fontSizePt) * fontSizePt * 1.2;
  }
  if (node.type === 'list') return estimateListHeightLines(node, maxWidthPt, fontPt) * fontPt * 1.2;
  if (node.type === 'code') return Math.max(node.value.split('\n').length, 1) * fontPt * 1.2;
  return estimateBlockLineCount(node, maxWidthPt, fontPt) * fontPt * 1.2;
}

// Each `<li>`'s own first paragraph can wrap to more than one line once
// squeezed into a real column width (the common case this was missing: a
// multi-column layout row — see slide.js's renderLayoutBlock — halves
// maxWidthPt for each cell, so a bullet like "Bank : <long bank name>" that
// fits on one line at full width often doesn't at half width). Subtracting
// the real bullet hang indent (BULLET_BASE_MARL, the same constant
// renderListParagraphs itself uses for marL) before wrapping matches how
// much horizontal room a bullet's TEXT actually gets, not the box's full
// width. Nested sub-lists/other block children after the item's first
// paragraph (see renderListParagraphs' own `rest` handling) still aren't
// modeled here — same "not a full layout engine" scope limit as elsewhere —
// only the single most common case (one wrapped text line per bullet) is
// fixed.
function estimateListHeightLines(node, maxWidthPt, fontSizePt, lineCounter = estimateBlockLineCount) {
  const items = node.items ?? [];
  if (!items.length) return 1;
  const indentPt = BULLET_BASE_MARL / EMU_PER_PT;
  const itemMaxWidthPt = maxWidthPt != null ? Math.max(1, maxWidthPt - indentPt) : undefined;
  return items.reduce((sum, item) => {
    const first = item.children?.[0];
    return sum + (first ? lineCounter(first, itemMaxWidthPt, fontSizePt) : 1);
  }, 0);
}

function estimateBlockLineCount(node, maxWidthPt, fontSizePt) {
  if (maxWidthPt == null) return 1;
  const text = flattenInline(node.children ?? []);
  if (!text) return 1;
  return Math.max(1, estimateWrappedLines(text, maxWidthPt, fontSizePt, resolveOwnFontFamily(node)).length);
}

// estimateWrappedLines (font-metrics.js) deliberately never splits a single
// unbroken token wider than maxWidthPt onto more than one of ITS OWN lines
// (see that function's own tested contract — a long word/badge label stays
// intact rather than being character-split). A real renderer still visually
// wraps such a token across several rows once it doesn't fit, so a TABLE
// CELL holding one — this project's own "Image Source" column: a bare URL
// or Windows file path directly inside a `<td>`, no spaces anywhere in it —
// needs that overflow accounted for here, on top of whatever ordinary
// word-level wrapping already produced, or the row is sized for far fewer
// lines than it actually needs once real PowerPoint force-wraps it. Only
// used for table-cell height estimation (rowNaturalHeightEmu below) — every
// OTHER caller of estimateBlockLineCount (card/badge sizing) keeps the
// original, deliberately-word-only behavior.
function estimateCellLineCount(node, maxWidthPt, fontSizePt) {
  if (maxWidthPt == null) return 1;
  const text = flattenInline(node.children ?? []);
  if (!text) return 1;
  const fontFamily = resolveOwnFontFamily(node);
  const lines = estimateWrappedLines(text, maxWidthPt, fontSizePt, fontFamily);
  return lines.reduce((sum, line) => {
    const lineWidthPt = estimateTextWidth(line, fontSizePt, fontFamily);
    return sum + Math.max(1, Math.ceil(lineWidthPt / maxWidthPt));
  }, 0);
}

// The card's own natural height in EMU — a background rect/roundRect sized
// to fit `nodes`' combined estimated content height plus padding. Exported
// so pptx/slide.js's renderLayoutBlock can compute a uniform row height
// across several cells *before* rendering any of them (each cell in a flex/
// grid row needs to end up the same height, not just its own content's).
// `maxWidthEmu` is the shape's own full outer width (before padding is
// subtracted) — passed by both call sites now that the column/box width is
// always known at the point they call this.
export function estimateCardHeightEmu(nodes, padding = DEFAULT_CARD_PADDING_PT, maxWidthEmu, rctx) {
  const padTopPt = padding.top ?? DEFAULT_CARD_PADDING_PT.top;
  const padBottomPt = padding.bottom ?? DEFAULT_CARD_PADDING_PT.bottom;
  const padLeftPt = padding.left ?? DEFAULT_CARD_PADDING_PT.left;
  const padRightPt = padding.right ?? DEFAULT_CARD_PADDING_PT.right;
  const maxWidthPt = maxWidthEmu != null ? Math.max(1, maxWidthEmu / EMU_PER_PT - padLeftPt - padRightPt) : undefined;
  const contentHeightPt = nodes.reduce((sum, n) => sum + estimateCardBlockHeightPt(n, maxWidthPt, rctx) + CARD_BLOCK_SPACING_PT, 0);
  return Math.round(contentHeightPt * EMU_PER_PT) + padTopPt * EMU_PER_PT + padBottomPt * EMU_PER_PT;
}

// A per-side border (`border-left`/`border-top`/etc — see dom-to-idm.js's
// extractBlockSideBorder, the common blockquote/accent-stripe pattern) has no
// direct DrawingML equivalent on the card's own shape: `a:ln` is one uniform
// outline around all 4 sides, with no "just this edge" option the way Word's
// `w:pBdr` has per-edge entries. Faking it as a second, thin solid-filled
// rectangle sitting flush against the card's edge is the same trick real
// slide software uses for an accent bar — currently only wired up for the
// left edge (the one real-world case seen so far, e.g. a blockquote), not a
// generic all-4-edges mechanism.
function renderAccentBarXml(border, box, cy, shapeId) {
  const widthEmu = Math.round((border.sizeEighthPt / 8) * EMU_PER_PT);
  const b = new XmlBuilder();
  b.el('p:sp');
  b.el('p:nvSpPr');
  b.self('p:cNvPr', { id: shapeId, name: `Accent ${shapeId}` });
  b.self('p:cNvSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x: box.x, y: box.y });
  b.self('a:ext', { cx: widthEmu, cy });
  b.up();
  b.el('a:prstGeom', { prst: 'rect' });
  b.self('a:avLst');
  b.up();
  b.el('a:solidFill');
  b.self('a:srgbClr', { val: border.color });
  b.up();
  b.up(); // p:spPr
  b.el('p:txBody');
  b.self('a:bodyPr');
  b.self('a:lstStyle');
  b.up();
  b.up();
  return b.end();
}

// Returns `{ xml, cy }` for one `p:sp` shape holding every node in `nodes`
// (all sharing one `cardGroup`) — a background rect/roundRect sized to fit
// their combined estimated height, with their own paragraphs rendered
// inside it. `nodes[0]` carries the group's border/fill/padding (mapElement
// stamps the same values onto every member), so only it is read for those.
// `forcedCy`, when given, overrides the natural content-fit height — used by
// renderLayoutBlock so every cell in one flex/grid row ends up the same
// height instead of each shrink-fitting to its own (possibly shorter) text.
// `accentShapeId`, when given (and `nodes[0]` has its own `borderLeft` — see
// renderAccentBarXml above), draws a second thin shape as that left-edge
// accent stripe; omitted entirely when there's no borderLeft to draw, or when
// the caller (renderLayoutBlock's flex/grid cells, which don't currently
// budget a second shape id per cell) didn't provide one.
export function renderCardShape(nodes, rctx, box, shapeId, { forcedCy, accentShapeId } = {}) {
  const first = nodes[0];
  const padding = first.padding ?? DEFAULT_CARD_PADDING_PT;
  const padLeft = (padding.left ?? DEFAULT_CARD_PADDING_PT.left) * EMU_PER_PT;
  const padRight = (padding.right ?? DEFAULT_CARD_PADDING_PT.right) * EMU_PER_PT;
  const padTop = (padding.top ?? DEFAULT_CARD_PADDING_PT.top) * EMU_PER_PT;
  const padBottom = (padding.bottom ?? DEFAULT_CARD_PADDING_PT.bottom) * EMU_PER_PT;

  // An explicit CSS `height` on the source container (e.g. a full-height
  // `.hero{height:900px}` banner) is a real authored size — same "explicit
  // beats inferred" precedence renderStyledBoxShape's own widthPx/heightPx
  // handling already uses for a badge — so it wins over the rough per-block
  // estimate below, never shrinking under it so real wrapped text is never
  // clipped by an under-estimated explicit height. Still capped to `box.cy`
  // (the slide's own remaining space), same as the plain estimate always was.
  const contentCy = estimateCardHeightEmu(nodes, padding, box.cx, rctx);
  const naturalCy = first.heightPx != null ? Math.max(Math.round(first.heightPx * EMU_PER_PX), contentCy) : contentCy;
  const cx = box.cx;
  const cy = Math.max(1, forcedCy ?? Math.min(box.cy, naturalCy));

  const b = new XmlBuilder();
  b.el('p:sp');
  b.el('p:nvSpPr');
  b.self('p:cNvPr', { id: shapeId, name: `Card ${shapeId}` });
  b.self('p:cNvSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x: box.x, y: box.y });
  b.self('a:ext', { cx, cy });
  b.up();
  // A real CSS `border-radius` (see hasBorderRadius) also earns a rounded
  // shape even with no `border` at all — e.g. `.gradient{border-radius:12px;
  // background:linear-gradient(...)}`, a borderless gradient panel that
  // still needs rounded corners like any other rounded box. `circular`
  // (border-radius:50%+) wins over both — see cardPresetGeometry.
  b.el('a:prstGeom', { prst: cardPresetGeometry(first) });
  b.self('a:avLst');
  b.up();
  b.raw(shapeFillXml(first));
  b.raw(shapeLineXml(first.border));
  if (first.shadow) b.raw(shadowEffectXml());
  b.up(); // p:spPr
  b.el('p:txBody');
  b.el('a:bodyPr', {
    wrap: 'square', anchor: 't',
    lIns: Math.round(padLeft), tIns: Math.round(padTop), rIns: Math.round(padRight), bIns: Math.round(padBottom),
  });
  // Our own height is an estimate, not a measurement — spAutoFit tells
  // PowerPoint to resize the shape to its actual text when the file is
  // opened, so a wrong guess self-corrects instead of clipping or leaving a
  // gap (see font-metrics.js's own module doc for why no exact measurement
  // is possible here).
  b.self('a:spAutoFit');
  b.up();
  b.self('a:lstStyle');
  // Each member's own backgroundColor is already the card shape's own fill
  // above (all members share it, per mapElement's stamping) — same
  // redundant-highlight reasoning renderStyledBoxShape's own strip already
  // documents.
  b.raw(nodes.flatMap((n) => renderBodyBlockParagraphs({ ...n, backgroundColor: undefined }, rctx)).join(''));
  b.up(); // p:txBody
  b.up(); // p:sp
  const accentXml = first.borderLeft && accentShapeId != null ? renderAccentBarXml(first.borderLeft, box, cy, accentShapeId) : '';
  return { xml: b.end() + accentXml, cy };
}

/* ─── Tables ─────────────────────────────────────────────────────────────── */

// A `<a:tr h="...">` is only a MINIMUM row height, not a fixed one — if the
// actual text at open-time doesn't fit within it, PowerPoint/LibreOffice grow
// the row to fit, silently making the real rendered table taller than
// whatever cy this file told slide.js's cursor to advance by. That mismatch
// is what used to make a shorter follow-up shape (e.g. this project's own
// totals table right after the line-items table) get positioned assuming the
// declared — not the real, auto-grown — height, overlapping the table above
// it. Estimating a natural per-row height up front (rather than an arbitrary
// flat constant) keeps the declared and real heights in sync. No `a:rPr sz`
// on a run (the common case here — see realTableCellXml) falls back to
// ECMA-376's own default of 1800 (18pt), not a guess — that's the literal
// spec default `CT_TextCharacterProperties/a:rPr` renders at when `sz` is
// omitted, matching this file's own DEFAULT_CARD_FONT_PT used everywhere
// else a DrawingML text body has no explicit size.
const DEFAULT_CELL_FONT_PT = 18;
// ECMA-376's own default `a:tcPr` cell insets when none are declared here
// (this file never sets lIns/tIns/rIns/bIns) — marT + marB, the only two
// that affect row height (marL/marR only affect column width headroom).
const TABLE_CELL_VPAD_EMU = 91440; // 45720 (marT) + 45720 (marB)
// The horizontal counterpart — marL + marR — narrows the width a cell's own
// text actually wraps WITHIN, below the column's full declared width (see
// cellContentHeightPt's own maxWidthPt param). Without this, a long line
// (e.g. this project's own "Image Source" column, holding a bare URL or
// Windows file path with no <p> wrapper) got measured against the FULL
// column width instead of the narrower area text really flows into, under-
// counting how many lines it actually wraps to.
const TABLE_CELL_HPAD_EMU = 182880; // 91440 (marL) + 91440 (marR)

// One block child's own rendered height, in points — mirrors
// realTableCellXml's own block-vs-inline split (its `children.every(
// isBlockNode)` check) AND renderBodyBlockParagraphs' own per-type font size
// (a heading renders at DEFAULT_HEADING_SIZE_PT, everything else at the
// cell's own base size), so this estimate stays in sync with what actually
// gets drawn. `baseFontPt` is the cell's own explicit fontSizePt (scaled,
// matching realTableCellXml) or the spec default when it declares none.
function blockContentHeightPt(node, baseFontPt, maxWidthPt) {
  if (node.type === 'heading') {
    const fontSizePt = DEFAULT_HEADING_SIZE_PT[node.level] ?? baseFontPt;
    return estimateCellLineCount(node, maxWidthPt, fontSizePt) * fontSizePt * 1.2;
  }
  if (node.type === 'list') return estimateListHeightLines(node, maxWidthPt, baseFontPt, estimateCellLineCount) * baseFontPt * 1.2;
  if (node.type === 'code') return Math.max(node.value.split('\n').length, 1) * baseFontPt * 1.2;
  // paragraph/quote/section — real wrapped line count (see
  // estimateCellLineCount), not always exactly one regardless of how long
  // the line actually is.
  return estimateCellLineCount(node, maxWidthPt, baseFontPt) * baseFontPt * 1.2;
}

// A cell's own total content height, in points — the sum of every block's
// own height (see blockContentHeightPt) for a block-content cell (e.g. this
// project's own "Payment Details" cell: an `<h3>` label at its own bigger
// heading size, followed by a `<ul>` of 4 plain-size items — all sharing one
// background, so band-splitting — splitIntoBackgroundBands — never
// separates them into their own rows; or, after band-splitting DID apply, a
// band cell holding just ONE heading child on its own — e.g. this project's
// own "From" band row), or just one line for a plain inline cell (the
// overwhelming common case). The `children.length` check (>=1, not >1)
// deliberately mirrors realTableCellXml's OWN identical condition for
// switching to its block-rendering path — a single-heading band cell must
// use ITS block branch here too, or this estimate would assume a single
// PLAIN line and under-count a heading's own bigger font size. Without
// summing every block's own height, a row was sized for ONE line while a
// cell like "Payment Details" actually needed five, visibly overflowing
// into whatever came right after the table.
function cellContentHeightPt(cell, rctx, maxWidthPt) {
  const baseFontPt = cell.fontSizePt != null ? cell.fontSizePt * CSS_TO_SLIDE_FONT_SCALE : resolveDefaultFontPt(rctx, DEFAULT_CELL_FONT_PT);
  const children = cell.children ?? [];
  if (children.length && children.every(isBlockNode)) {
    return children.reduce((sum, child) => sum + blockContentHeightPt(child, baseFontPt, maxWidthPt), 0);
  }
  if (!children.length) return baseFontPt * 1.2;
  // A plain inline cell (no <p> wrapper — e.g. this project's own "Image
  // Source" column: a bare URL/file path directly inside a <td>) still
  // wraps to real multiple lines once it's longer than the column can hold
  // in one line — estimateCellLineCount measures that against the cell's
  // OWN content, not always the flat single-line guess this used to return
  // regardless of length.
  return estimateCellLineCount(cell, maxWidthPt, baseFontPt) * baseFontPt * 1.2;
}

// The tallest cell's own total content height in `gridRow` (see
// cellContentHeightPt) — a ROW's height is governed by whichever cell in it
// needs the most room, matching how a browser sizes a real `<tr>`.
// `colWidthEmu` is the table's own uniform per-column width (renderTableShape
// gives every column the same width — see its own colWidth) — a colspan
// slot gets the sum of every column it covers, same as any other exporter's
// colspan-aware width computation.
// A cell's own `<img>` needs its OWN real display height accounted for
// here too — this estimate previously only ever looked at TEXT content
// (cellContentHeightPt), so an image-only "Preview"-style cell always came
// back as a single text line's height (~18-24pt) regardless of the image's
// own real/declared size, squeezing every photo down into a sliver far
// shorter than a browser would ever render that same `<td><img
// style="width:80px;height:80px"></td>` at. Mirrors imageNaturalSize/
// fitWithin's own "shrink to fit the available width, never upscale"
// aspect-fit logic (see renderPictureShape), just measured against the
// CELL's own available width instead of a whole slide's content box.
function cellImageNaturalHeightEmu(cell, rctx, availableWidthEmu) {
  const imageNode = findCellImageNode(cell);
  if (!imageNode) return 0;
  const assetInfo = rctx.assetInfoByAssetId?.get(imageNode.assetId);
  if (!assetInfo) return 0;
  const natural = imageNaturalSize(imageNode, assetInfo);
  const scale = Math.min(1, availableWidthEmu / natural.cx);
  return Math.round(natural.cy * scale) + TABLE_CELL_VPAD_EMU;
}

function rowNaturalHeightEmu(gridRow, rctx, colWidthEmu) {
  let maxHeightPt = resolveDefaultFontPt(rctx, DEFAULT_CELL_FONT_PT) * 1.2;
  let maxImageHeightEmu = 0;
  for (const slot of gridRow) {
    if (slot.continuation || !slot.cell) continue;
    const spanWidthEmu = colWidthEmu * (slot.colspan ?? 1);
    const availableWidthEmu = Math.max(1, spanWidthEmu - TABLE_CELL_HPAD_EMU);
    const maxWidthPt = availableWidthEmu / EMU_PER_PT;
    maxHeightPt = Math.max(maxHeightPt, cellContentHeightPt(slot.cell, rctx, maxWidthPt));
    maxImageHeightEmu = Math.max(maxImageHeightEmu, cellImageNaturalHeightEmu(slot.cell, rctx, availableWidthEmu));
  }
  const textHeightEmu = Math.round(maxHeightPt * EMU_PER_PT) + TABLE_CELL_VPAD_EMU;
  return Math.max(textHeightEmu, maxImageHeightEmu);
}

// A plain default border on every cell — unlike DOCX's TableGrid style,
// DrawingML tables render with no visible lines at all unless every cell's
// own tcPr declares one explicitly. Only used as a fallback for a
// genuinely bare, zero-CSS table (see tableHasExplicitStyling) — a table
// the author styled in some way, even choosing deliberately no border
// (this project's own totals `<table style="width:50%;margin-left:auto">`,
// with per-cell `text-align:right` but no border), instead draws only
// whichever border it actually declared (see cellBorderLinesXml).
function cellBorderLines() {
  const b = new XmlBuilder();
  for (const tag of ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']) {
    b.el(tag, { w: 6350 });
    b.el('a:solidFill');
    b.self('a:srgbClr', { val: '999999' });
    b.up();
    b.up();
  }
  return b.end();
}

// Mirrors PDF's blocks.js `CELL_STYLE_FIELDS`/`cellHasExplicitStyle`/
// `tableHasExplicitStyling` (itself mirroring DOCX's own identical
// heuristic) — see cellBorderLines' own note for the full reasoning.
const CELL_STYLE_FIELDS = [
  'backgroundColor', 'borderTop', 'borderRight', 'borderBottom', 'borderLeft', 'noBorder',
  'padding', 'verticalAlign', 'color', 'bold', 'italic', 'underline', 'strikethrough',
  'fontSizePt', 'direction', 'writingMode', 'rotateDeg',
];

function cellHasExplicitStyle(cell) {
  if (!cell) return false;
  if (CELL_STYLE_FIELDS.some((field) => cell[field] !== undefined)) return true;
  return cell.align !== undefined && !(cell.header && cell.align === 'center');
}

function tableHasExplicitStyling(table, grid) {
  if (table.border != null || table.widthPercent != null || table.marginAuto != null) return true;
  for (const row of grid) {
    for (const slot of row) {
      if (!slot.continuation && cellHasExplicitStyle(slot.cell)) return true;
    }
  }
  return false;
}

// One declared side's border for an explicitly-styled table — the cell's
// own if present, else the table's own uniform `border` as every cell's
// shared default for that side, never applied over an explicit `noBorder`.
function resolveCellEdgeBorder(cell, field, tableBorder) {
  const own = cell?.[field];
  if (own) return own;
  return cell?.noBorder ? undefined : tableBorder;
}

const EDGE_BORDER_TAGS = [['a:lnL', 'borderLeft'], ['a:lnR', 'borderRight'], ['a:lnT', 'borderTop'], ['a:lnB', 'borderBottom']];

// Draws whichever sides `cell` (or the table's own uniform border) actually
// declare a real line, and an EXPLICIT `<a:noFill/>` for every other side —
// the explicitly-styled-table counterpart to cellBorderLines' unconditional
// flat default. Omitting a side's `<a:ln*>` element entirely (rather than
// giving it a real, if empty, line-properties child) does NOT mean "no
// border" in DrawingML the way it would in a from-scratch renderer — it
// means "inherit whatever this table's own referenced/default table STYLE
// says for that side," and PowerPoint's own built-in default style (applied
// whenever `<a:tblPr>` carries no `<a:tableStyleId>`, exactly what this
// project's tables emit) draws a visible bordered/banded grid. An explicit
// `<a:noFill/>` line, by contrast, is a real per-cell override that beats
// the style default the same way an explicit color does — the only way to
// make an intentionally borderless table actually render borderless.
// A cell often leaves one side (`borderBottom`, say) undeclared simply
// because that particular `<td>` never mentioned it — not because the
// author explicitly wants no line there — and resolveCellEdgeBorder's own
// fallback for that case is a plain `<a:noFill/>`. When the ADJACENT cell
// across that exact same shared edge (the row above's own bottom, for a
// cell's own top; the row below's own top, for a cell's own bottom) DOES
// declare a real border, at least one real-world OOXML consumer was
// observed to let that undeclared/noFill side visually win at the shared
// boundary, hiding a border that (from the XML alone) looks perfectly
// declared — e.g. this project's own totals table: Tax's cells never
// mention a bottom border, Total's cells right below explicitly declare
// `border-top`, and the line between them silently failed to render even
// after independently confirming the XML carried a correctly-sized,
// correctly-colored `<a:lnT>`. Mirroring a declared border onto its
// undeclared neighbor — so BOTH sides of a shared edge agree it's a real
// line, not one real + one implicit-noFill — removes that ambiguity for
// every renderer, not just the one known to mishandle it. `neighbor` is
// only ever given for the two vertical sides (top/bottom); left/right
// mirroring across colspan boundaries isn't wired up here — not the
// reported problem, and a full colspan-aware version adds real complexity
// for a case with no known failure yet.
function cellDeclaredBorderLines(cell, tableBorder, neighbor = {}) {
  const b = new XmlBuilder();
  for (const [tag, field] of EDGE_BORDER_TAGS) {
    let border = resolveCellEdgeBorder(cell, field, tableBorder);
    if (!border && !cell?.noBorder) {
      if (field === 'borderTop') border = neighbor.aboveCell?.noBorder ? undefined : neighbor.aboveCell?.borderBottom;
      else if (field === 'borderBottom') border = neighbor.belowCell?.noBorder ? undefined : neighbor.belowCell?.borderTop;
    }
    if (border) {
      b.el(tag, { w: Math.round((border.sizeEighthPt / 8) * EMU_PER_PT) });
      b.el('a:solidFill');
      b.self('a:srgbClr', { val: border.color });
      b.up();
      b.up();
    } else {
      b.el(tag);
      b.self('a:noFill');
      b.up();
    }
  }
  return b.end();
}

// Unlike WordprocessingML (where a single <w:tc gridSpan="N"/> covers N grid
// columns on its own), DrawingML requires one <a:tc> per grid column even
// inside a merge — a colspan needs N-1 extra `hMerge="1"` placeholder cells
// after the real one, and a rowspan's covered rows each need a `vMerge="1"`
// placeholder (see ECMA-376 §21.1.3.8/.16) — so a slot with colspan>1 expands
// into `colspan` sibling <a:tc> elements, not one.
function tableRowSlotXml(slot, isHeaderRow, borderMode, rctx, neighbor) {
  const colspan = slot.colspan ?? 1;
  if (slot.continuation) {
    // originCell carries the merged region's own styling (background, in
    // particular) — without threading it through, only the origin's own
    // grid position would show the fill, leaving the rest of a colored
    // merged region visibly unfilled.
    const backgroundColor = slot.originCell?.backgroundColor;
    let out = '';
    for (let i = 0; i < colspan; i++) out += mergePlaceholderCell({ vMerge: true, hMerge: i > 0, backgroundColor }, slot.originCell, borderMode);
    return out;
  }
  const rowspan = slot.cell.rowspan ?? 1;
  let out = realTableCellXml(slot.cell, colspan, rowspan, isHeaderRow, borderMode, rctx, neighbor);
  for (let i = 1; i < colspan; i++) out += mergePlaceholderCell({ hMerge: true, backgroundColor: slot.cell.backgroundColor }, slot.cell, borderMode);
  return out;
}

// CT_TableCellProperties is a strict sequence: border lines first, then the
// fill — the reverse order silently fails to apply the fill in PowerPoint.
function tcPrXml(backgroundColor, cell, borderMode, neighbor) {
  const b = new XmlBuilder();
  b.el('a:tcPr');
  b.raw(borderMode.defaultGrid ? cellBorderLines() : cellDeclaredBorderLines(cell, borderMode.tableBorder, neighbor));
  if (backgroundColor) {
    b.el('a:solidFill');
    b.self('a:srgbClr', { val: backgroundColor });
    b.up();
  }
  b.up();
  return b.end();
}

// A cell's own `<img>` (bare, or a `<figure>`'s single image — its caption
// has nowhere to go once the image is drawn as its own floating shape
// rather than flowed content, so it's dropped in this pass, same scope
// limit as everywhere else images-in-cells weren't handled at all) — only
// the FIRST one found, since only one picture is drawn per cell; a cell
// with several images has no representation this mechanism can give it.
// Doesn't recurse into a nested `table` — an image inside a NESTED table
// belongs to that inner table's own cell, not this outer cell's own
// picture (see renderTableShape, which walks each cell for this exact
// reason: a native `a:tc` can't nest a real `p:pic`, so the picture is
// instead drawn as its own floating shape positioned directly over the
// cell's own real bounds, once the table's own row/column geometry is
// known — aspect-fit and centered within it, same as any other image,
// never the non-uniform stretch-to-cover-the-whole-cell a `<a:blipFill>`
// would have forced).
function findCellImageNode(cell) {
  const walk = (nodes) => {
    for (const node of nodes ?? []) {
      if (!node) continue;
      if (node.type === 'image') return node;
      if (node.type === 'figure') return node.image;
      if (node.type === 'table') continue;
      const found = (node.children && walk(node.children)) || (node.items && walk(node.items.flatMap((i) => i.children)));
      if (found) return found;
    }
    return undefined;
  };
  return walk(cell?.children);
}

function mergePlaceholderCell({ vMerge, hMerge, backgroundColor }, cell, borderMode) {
  const b = new XmlBuilder();
  b.el('a:tc', { hMerge: hMerge ? 1 : undefined, vMerge: vMerge ? 1 : undefined });
  b.el('a:txBody');
  b.self('a:bodyPr');
  b.self('a:lstStyle');
  b.el('a:p').up();
  b.up();
  b.raw(tcPrXml(backgroundColor, cell, borderMode));
  b.up();
  return b.end();
}

// Renders a cell's own actual content — preserving a <strong>/<em>/<u>/<a>
// child's own formatting via renderInlineRuns, the same way a plain
// paragraph's text renders elsewhere in this file — rather than flattening
// it down to one plain-text run first (flattenCellText, still used by CSV/
// TXT/the content-based column-width heuristics that only ever need a bare
// string). Flattening first was silently dropping every inline `<b>`/`<i>`/
// `<u>`/`<a>` inside a `<td>` — e.g. this project's own totals table's
// `<td><b>Total</b></td>` rendered as plain, unbolded text — since a cell's
// own CSS-level `bold`/`italic`/etc (read once per cell by mapCells, from the
// cell/row's own `style` attribute only) is a completely different signal
// from an inline child tag's formatting, and only the former ever reached
// this function before.
function realTableCellXml(cell, colspan, rowspan, isHeaderRow, borderMode, rctx, neighbor) {
  const algn = ALIGN_MAP[cell.align];
  const children = cell.children ?? [];
  // The cell's own CSS-level defaults (a `<td style="font-weight:bold">`, or
  // `isHeaderRow`'s implicit bold) — each inline child still layers its own
  // formatting on top of these via renderRun, exactly like a plain
  // paragraph's renderInlineRuns call does, so neither signal shadows the
  // other.
  const baseProps = {
    bold: isHeaderRow || cell.bold, italic: cell.italic, underline: cell.underline,
    strikethrough: cell.strikethrough, color: cell.color,
    // An explicit cell fontSizePt always wins; absent that, the document's
    // own body font-size (rctx.bodyFontSizePt — see resolveDefaultFontPt)
    // still beats leaving `sz` off entirely, which would otherwise fall
    // back to OOXML's own spec default (1800=18pt) regardless of what the
    // source page actually declared.
    fontSizePt: cell.fontSizePt != null ? cell.fontSizePt * CSS_TO_SLIDE_FONT_SCALE : resolveDefaultFontPt(rctx, undefined),
  };
  const b = new XmlBuilder();
  b.el('a:tc', { gridSpan: colspan > 1 ? colspan : undefined, rowSpan: rowspan > 1 ? rowspan : undefined });
  b.el('a:txBody');
  b.self('a:bodyPr');
  b.self('a:lstStyle');
  // A cell with block content (a nested <p>/<ul>/<table>/... — see
  // mapCells' own hasBlockContent check) needs its own real paragraphs, the
  // same block-vs-inline split flattenCellText already draws for its plain-
  // text callers, just rendered as real content instead of joined text.
  if (children.length && children.every(isBlockNode)) {
    b.raw(children.flatMap((child) => renderBodyBlockParagraphs(child, rctx)).join(''));
  } else {
    b.el('a:p');
    if (algn) { b.el('a:pPr', { algn }); b.up(); }
    if (children.length) b.raw(renderInlineRuns(children, rctx, baseProps));
    b.up();
  }
  b.up();
  b.raw(tcPrXml(cell.backgroundColor, cell, borderMode, neighbor));
  b.up();
  return b.end();
}

// Renders a TableNode as its own `<p:graphicFrame>`/`<a:tbl>` shape, clamped
// to fit within `box` — equal column widths and a fixed row height (no
// content-based sizing, same simplification as the rest of this v1 pass).
// A <td> holding several distinctly-background-colored block children (the
// common "colored label bar followed by plain address lines" pattern — e.g.
// `<td><h3 style="background:#1F4E79;color:white">From</h3><p>...</p>
// <p>...</p></td>`) can't have that per-block background rendered within
// ONE `<a:tc>` — a cell's own fill is all-or-nothing for its whole area, the
// same way it is in Excel; there's no per-paragraph shading in DrawingML
// tables either. Splitting the row into a couple of dedicated physical rows
// (the heading gets its OWN row/cell with its OWN real `<a:tcPr>` fill, the
// plain lines get their own row below with none) is what actually gives
// "From"/"Bill To" a real full-cell-width colored bar matching the browser,
// rather than the tight run-highlight approximation renderBodyBlockParagraphs'
// own `bg` handling falls back to for a plain single-block heading (fine for
// a short badge, but never spans a whole cell's width the way a genuine
// `<a:tcPr>` fill does). Mirrors XLSX's own identical
// splitIntoBackgroundBands/expandRowForBands (worksheet.js) — same IDM
// shapes, same reasoning, just building PPTX rows/rowSpan instead of
// SpreadsheetML rows/mergeCells.
function splitIntoBackgroundBands(cell) {
  const children = cell.children ?? [];
  if (children.length < 2 || !children.every(isBlockNode)) return null;
  const bands = [];
  for (const child of children) {
    const backgroundColor = child.backgroundColor ?? cell.backgroundColor;
    const last = bands[bands.length - 1];
    if (last && last.backgroundColor === backgroundColor) last.children.push(child);
    else bands.push({ backgroundColor, children: [child] });
  }
  return bands.length > 1 ? bands : null;
}

// Expands one TableRowNode into several physical rows when any of its cells
// has multiple background bands (see splitIntoBackgroundBands above). A
// sibling cell that DOESN'T split (the common case: a plain data cell, or
// one with only a single background throughout) spans every physical row
// this logical row expanded into via a real rowSpan (expandTableGrid, from
// table-grid.js, already knows how to turn that into `vMerge` continuation
// slots — no changes needed there), so it still reads as one continuous box
// rather than being cut into empty-looking fragments. A cell with FEWER
// bands than this row's max only pads its OWN *last* band with that
// rowSpan (the same "slack absorbed at the end" convention already used
// throughout this project). border-top/-bottom are only kept on the
// first/last band respectively (left/right stay on every band, since
// column edges don't change when a row splits vertically) — otherwise every
// internal band boundary would draw a redundant double border line the
// original single-box design never had. Padding's top/bottom split the same
// way, so the visual gap doesn't double up at the fold between two bands.
function expandRowForBands(row) {
  const perCellBands = row.cells.map((cell) => ({ cell, bands: splitIntoBackgroundBands(cell) }));
  const maxBands = Math.max(1, ...perCellBands.map(({ bands }) => bands?.length ?? 1));
  if (maxBands === 1) return [row];

  const physicalRows = [];
  for (let bandIndex = 0; bandIndex < maxBands; bandIndex++) {
    const cells = [];
    for (const { cell, bands } of perCellBands) {
      if (!bands) {
        if (bandIndex === 0) cells.push({ ...cell, rowspan: maxBands });
        continue;
      }
      if (bandIndex >= bands.length) continue;
      const band = bands[bandIndex];
      const isFirstBand = bandIndex === 0;
      const isLastBand = bandIndex === bands.length - 1;
      const span = isLastBand ? maxBands - bandIndex : 1;
      cells.push({
        ...cell,
        backgroundColor: band.backgroundColor,
        children: band.children,
        borderTop: isFirstBand ? cell.borderTop : undefined,
        borderBottom: isLastBand ? cell.borderBottom : undefined,
        padding: cell.padding && {
          ...cell.padding,
          top: isFirstBand ? cell.padding.top : 0,
          bottom: isLastBand ? cell.padding.bottom : 0,
        },
        rowspan: span > 1 ? span : undefined,
      });
    }
    // Only the first physical row can plausibly inherit the source <tr>'s
    // own explicit CSS height — the rest are new rows this expansion itself
    // introduced, with no explicit height of their own to carry.
    physicalRows.push({ ...row, cells, heightPx: bandIndex === 0 ? row.heightPx : undefined });
  }
  return physicalRows;
}

function expandRowsForBands(rows) {
  return rows.flatMap(expandRowForBands);
}

export function renderTableShape(table, rctx, box, shapeId, nextShapeId) {
  const headRows = expandRowsForBands(table.head ?? []);
  const bodyRows = expandRowsForBands(table.rows);
  const allRows = [...headRows, ...bodyRows];
  if (allRows.length === 0) return { xml: '', cy: 0 };

  const { grid, columnCount } = expandTableGrid(allRows);

  // A table's own `width:N%` (e.g. this project's own "totals" summary
  // table: `width:50%;margin-left:auto`) narrows the drawn table instead of
  // always stretching it across the full content width, matching a
  // browser's own <table> sizing and the same handling PDF's blocks.js
  // (renderTable) and DOCX's document.js already give it. `margin-left:auto`
  // (with no matching margin-right:auto) floats the narrowed table to the
  // right edge; `margin:auto` on both sides centers it.
  const tableWidth = table.widthPercent != null
    ? Math.max(1, Math.round(box.cx * (table.widthPercent / 100)))
    : box.cx;
  let tableX = box.x;
  if (table.marginAuto === 'both') tableX = box.x + Math.round((box.cx - tableWidth) / 2);
  else if (table.marginAuto === 'left') tableX = box.x + (box.cx - tableWidth);

  const colWidth = Math.max(1, Math.floor(tableWidth / columnCount));
  // A compact, content-sized row height per row (see rowNaturalHeightEmu) —
  // NOT box.cy/grid.length stretched to fill whatever vertical space happens
  // to be left on the slide (the previous behavior): a short 4-row totals
  // table sitting above a mostly-empty rest-of-slide would otherwise balloon
  // each row up to a full inch tall, reading as huge gaps between rows
  // instead of a normal compact table. The whole table's rows are only
  // scaled down together when the box genuinely doesn't have room for their
  // combined natural height — never grown to fill unused space.
  const naturalRowHeights = grid.map((gridRow) => rowNaturalHeightEmu(gridRow, rctx, colWidth));
  const naturalTotal = naturalRowHeights.reduce((a, h) => a + h, 0);
  const shrink = naturalTotal > box.cy ? box.cy / naturalTotal : 1;
  const rowHeights = naturalRowHeights.map((h) => Math.max(1, Math.round(h * shrink)));
  const totalHeight = rowHeights.reduce((a, h) => a + h, 0);
  // A genuinely bare, zero-CSS table still gets the flat default gray grid
  // every cell always drew before this — a table the author styled in some
  // way (even choosing deliberately no border) instead only ever gets
  // whichever border it actually declared (see tableHasExplicitStyling).
  const borderMode = tableHasExplicitStyling(table, grid)
    ? { defaultGrid: false, tableBorder: table.border }
    : { defaultGrid: true, tableBorder: undefined };

  return buildTableGraphicFrame(grid, columnCount, colWidth, rowHeights, totalHeight, { x: tableX, y: box.y }, shapeId, headRows.length, borderMode, rctx, nextShapeId);
}

// The resolved cell (continuation slots resolve to their originCell) sitting
// at the exact same [col, col+span) column range in `row` — used to find the
// vertically-adjacent cell to mirror a border against (see
// cellDeclaredBorderLines' own note). Only matches an EXACT span alignment;
// a colspan that doesn't line up 1:1 between the two rows is left
// unmirrored rather than guessing which of several neighbors to trust.
function findAlignedCell(row, col, span) {
  if (!row) return undefined;
  let c = 0;
  for (const slot of row) {
    const slotSpan = slot.colspan ?? 1;
    if (c === col && slotSpan === span) return slot.continuation ? slot.originCell : slot.cell;
    c += slotSpan;
  }
  return undefined;
}

function buildTableGraphicFrame(grid, columnCount, colWidth, rowHeights, totalHeight, pos, shapeId, headerRowCount, borderMode, rctx, nextShapeId) {
  const b = new XmlBuilder();
  b.el('p:graphicFrame');
  b.el('p:nvGraphicFramePr');
  b.self('p:cNvPr', { id: shapeId, name: `Table ${shapeId}` });
  b.self('p:cNvGraphicFramePr');
  b.self('p:nvPr');
  b.up();
  b.el('p:xfrm');
  b.self('a:off', { x: pos.x, y: pos.y });
  b.self('a:ext', { cx: colWidth * columnCount, cy: totalHeight });
  b.up();
  b.el('a:graphic');
  b.el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/table' });
  b.el('a:tbl');
  b.self('a:tblPr');
  b.el('a:tblGrid');
  for (let c = 0; c < columnCount; c++) b.self('a:gridCol', { w: colWidth });
  b.up();
  // A native `a:tc` can't nest a real `p:pic` (see findCellImageNode's own
  // doc comment) — every cell holding an image is instead recorded here, by
  // its own real absolute EMU bounds (now knowable, since colWidth/
  // rowHeights are already resolved above — and rowHeights now already
  // accounts for an image's own natural size too, see rowNaturalHeightEmu),
  // and drawn as its OWN floating picture shape immediately after the
  // table itself, positioned exactly over that cell. A rowspan'd cell is
  // only measured against its own ORIGIN row's height (continuation slots
  // carry no `cell` of their own to look an image up on, same as their
  // text content is never rendered twice either) — a minor under-count for
  // the rare rowspan+image combination, not worth the extra bookkeeping
  // this file's other estimates already don't attempt for equally rare cases.
  const cellImages = [];
  let rowY = pos.y;
  grid.forEach((gridRow, r) => {
    b.el('a:tr', { h: rowHeights[r] });
    const aboveRow = r > 0 ? grid[r - 1] : undefined;
    const belowRow = r < grid.length - 1 ? grid[r + 1] : undefined;
    let col = 0;
    for (const slot of gridRow) {
      const span = slot.colspan ?? 1;
      const neighbor = {
        aboveCell: findAlignedCell(aboveRow, col, span),
        belowCell: findAlignedCell(belowRow, col, span),
      };
      b.raw(tableRowSlotXml(slot, r < headerRowCount, borderMode, rctx, neighbor));
      if (!slot.continuation) {
        const cellImage = findCellImageNode(slot.cell);
        if (cellImage) {
          // Must match rowNaturalHeightEmu's own cellImageNaturalHeightEmu
          // call EXACTLY — same padded width, same padded height — or this
          // box and the row height it's drawn into disagree with each
          // other. Using the FULL (unpadded) column width here let
          // renderPictureShape's own fitWithin scale the image up slightly
          // past what the row height actually budgeted for (computed
          // against the NARROWER padded width), overflowing past the row's
          // own bottom edge in real PowerPoint.
          const cellWidthEmu = colWidth * span;
          const availableWidthEmu = Math.max(1, cellWidthEmu - TABLE_CELL_HPAD_EMU);
          const availableHeightEmu = Math.max(1, rowHeights[r] - TABLE_CELL_VPAD_EMU);
          cellImages.push({
            imageNode: cellImage,
            box: {
              x: pos.x + col * colWidth + Math.round((cellWidthEmu - availableWidthEmu) / 2),
              y: rowY + Math.round((rowHeights[r] - availableHeightEmu) / 2),
              cx: availableWidthEmu,
              cy: availableHeightEmu,
            },
          });
        }
      }
      col += span;
    }
    b.up();
    rowY += rowHeights[r];
  });
  b.up(); // a:tbl
  b.up(); // a:graphicData
  b.up(); // a:graphic
  b.up(); // p:graphicFrame

  let xml = b.end();
  for (const { imageNode, box } of cellImages) {
    xml += renderPictureShape(imageNode, rctx, box, nextShapeId(), 'middle').xml;
  }
  return { xml, cy: totalHeight };
}
