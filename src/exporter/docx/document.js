/**
 * word/document.xml — the document body. Built up incrementally:
 * W1 skeleton · W2 paragraphs · W3 headings/breaks · W4 inline formatting ·
 * W5 lists · W6 tables · W7 hyperlinks · W8 images · W9 metadata (elsewhere)
 * · W10 page setup from options.
 *
 * `rctx` (render context) carries the per-export managers that block/inline
 * renderers need — NumberingManager (W5), RelationshipManager (W7/W8) — so
 * document.xml and its .rels file can never drift out of sync. It's threaded
 * through every render function, not just the ones that currently use it,
 * so a hyperlink/image nested inside a list/table/quote still reaches it.
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';
import { REL_TYPES } from '../ooxml/package.js';
import { resolveFontFamilyName } from '../../fonts/resolver.js';
import { flattenCellText, flattenInline, isBlockNode } from '../../idm/flatten-text.js';
import { expandTableGrid } from '../shared/table-grid.js';
import { shapeFillXml, shapeLineXml, shadowEffectXml, needsDrawingShape, cardPresetGeometry, resolveSvgShapeGeometry } from '../shared/shape-style.js';
import { estimateShapeSize } from '../shared/font-metrics.js';
import { resolveImagePxSize } from '../shared/image-size.js';

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
};

// The four outer edges of a box border (paragraph, table, or cell) — the
// same order used everywhere a border/margin/padding is rendered per side.
const BOX_EDGES = ['top', 'left', 'bottom', 'right'];

export function buildDocumentXml(document, ctx, rctx) {
  const b = new XmlBuilder();
  b.el('w:document', {
    'xmlns:w': NS.w,
    'xmlns:r': NS.r,
    'xmlns:wp': NS.wp,
    'xmlns:a': NS.a,
    'xmlns:pic': NS.pic,
  });
  b.el('w:body');
  // The real page content width from convert()'s pageSize/orientation/margins
  // options (W10) — tables need this to size their columns against the
  // *actual* page, not a hardcoded assumption; a table's own width
  // computation happens during this renderBlocks call, so it must be set on
  // rctx before that, not derived only later when renderSectPr runs.
  rctx.pageContentWidthTwips = computeContentWidthTwips(ctx?.options);
  b.raw(renderBlocks(document.children, rctx));

  renderSectPr(b, ctx);

  return xmlDoc(b.end());
}

// ─── Block rendering ────────────────────────────────────────────────────────

// A `display:flex`/`display:grid` container's children (see
// dom-to-idm.js's mapLayoutBlocks) are tagged with a shared `layoutGroup` +
// `layoutItemIndex` — without special handling here they'd fall through to
// renderBlock one at a time and stack as plain vertical paragraphs, losing
// the side-by-side/grid arrangement entirely (e.g. a row of colored "Box 1/
// 2/3" divs reads as three separately-shaded lines, one under the other).
// Grouping consecutive same-`layoutGroup` nodes into one real Word table
// mirrors pptx/slide.js's own classifyContentUnits/renderLayoutBlock. Only
// applied at the top level (document.children) — a flex/grid container
// nested inside a list item or table cell is rare enough, and rendering a
// table-inside-a-table-cell/list-item correctly is involved enough, that it
// isn't worth the added risk for this pass; it still degrades to the same
// plain-paragraph-stack fallback as before this fix.
function renderBlocks(nodes, rctx) {
  const out = [];
  let i = 0;
  let lastWasTable = false;
  while (i < nodes.length) {
    const node = nodes[i];
    if ((node.type === 'paragraph' || node.type === 'heading') && node.layoutGroup !== undefined) {
      const layoutGroup = node.layoutGroup;
      const cells = [];
      let cell = [node];
      let j = i + 1;
      while (j < nodes.length && nodes[j].layoutGroup === layoutGroup) {
        if (nodes[j].layoutItemIndex !== cell[0].layoutItemIndex) {
          cells.push(cell);
          cell = [];
        }
        cell.push(nodes[j]);
        j++;
      }
      cells.push(cell);
      out.push(renderLayoutTable(cells, layoutGroup, rctx));
      i = j;
      lastWasTable = false;
      continue;
    }
    // A bordered/filled <div> wrapping several blocks (dom-to-idm.js's
    // mapElement stamps a shared `cardGroup` token on all of them — see
    // pptx/slide.js's identical grouping). Only promoted to a real
    // DrawingML shape (renderCardShapeDrawing) when at least one member
    // needs geometry plain w:pBdr/w:shd can't do at all (rounded corners,
    // a shadow — see needsDrawingShape); a plain flat-bordered card keeps
    // rendering each member as its own bordered/shaded paragraph, relying
    // on Word's own border-merging between adjacent identically-bordered
    // paragraphs to still read as one box, same as before this change.
    if ((node.type === 'paragraph' || node.type === 'heading') && node.cardGroup !== undefined) {
      const cardGroup = node.cardGroup;
      const group = [node];
      let j = i + 1;
      while (j < nodes.length && nodes[j].cardGroup === cardGroup) group.push(nodes[j++]);
      out.push(needsDrawingShape(node) ? renderCardShapeDrawing(group, rctx) : group.map((n) => renderBlock(n, rctx)).join(''));
      i = j;
      lastWasTable = false;
      continue;
    }
    // A standalone rounded/shadowed box with no cardGroup (a single <div>
    // that fully styled itself, e.g. `.gradient{border-radius:12px;
    // background:linear-gradient(...)}`) — same promotion, single-member
    // group. Excludes shrinkToFit (badges/pills): those get their own
    // dedicated dispatch inside renderBlock's `case 'paragraph'` instead
    // (renderShapedBadge for a circular/rounded one, renderShrinkToFitBox's
    // plain square-cornered table otherwise) — routing them through this
    // path too would just promote the same node to a shape twice.
    //
    // Requiring an actual border/backgroundGradient/shadow (NOT plain
    // backgroundColor alone) matters just as much here as cardGroupId's own
    // `border` requirement does above, and for the same reason
    // pptx/shapes.js's isStyledBoxNode requires it too: a page-level
    // wrapper's own plain background (`.slide{background:white;
    // border-radius:15px}`, no border of its own, so it never sets
    // cardGroup — see dom-to-idm.js's mapElement) bleeds `backgroundColor`
    // onto EVERY plain, otherwise-unstyled paragraph/heading inside it —
    // backgroundColor alone is exactly the "just happens to sit inside
    // something with a background" false positive, not a real bordered box
    // or gradient panel of its own. Without this, nearly every heading in an
    // otherwise plain document sitting inside such a wrapper would wrongly
    // become its own separate box. `shadow`, unlike backgroundColor, IS
    // trustworthy here: dom-to-idm.js's own shadowRoundedTrusted gate already
    // only ever lets it reach a node when that node is either the wrapper's
    // sole child or the wrapper also has a real border — the exact same
    // "genuinely one bounded panel" bar backgroundGradient/border already
    // clear, so a `.slide`-style many-children/no-border wrapper's shadow
    // never reaches here at all (see that gate's own note for the full
    // reasoning) and a real single-child `.shadow{background:white;
    // box-shadow:...}` div does.
    const isOwnBox = node.border || node.backgroundGradient || node.shadow;
    if ((node.type === 'paragraph' || node.type === 'heading') && !node.shrinkToFit && isOwnBox && needsDrawingShape(node)) {
      out.push(renderCardShapeDrawing([node], rctx));
      i++;
      lastWasTable = false;
      continue;
    }
    // Word silently merges two <w:tbl> elements into a single visual table
    // when it opens a document if there's nothing between them — any width/
    // alignment/border difference the second table declared (e.g. a
    // narrower, right-aligned totals table right after a full-width data
    // table) then gets swallowed into the first table's own shape. An empty
    // paragraph is the standard fix and matches how two separate <table>s
    // in the source HTML always render as two distinct tables, never one.
    if (node.type === 'table' && lastWasTable) out.push('<w:p/>');
    out.push(renderBlock(node, rctx));
    lastWasTable = node.type === 'table';
    i++;
  }
  return out.join('');
}

// A node's own border/background/padding — already rendered as this cell's
// own w:tcBorders/w:shd/w:tcMar below — would otherwise ALSO render as that
// same paragraph's own w:pBdr/w:shd/w:ind, doubling the box. Matches
// renderCardShape's own "strip it before rendering the paragraph" approach.
// `margin` is stripped too: it's the *group's own* outer spacing (e.g. a
// `.card{margin:20px 0}` meant to separate the whole card from whatever
// sits above/below it), stamped onto every member the same "if unset" way
// border/background are — left in place, it would add that same 20px
// before/after *each* member paragraph individually (a heading plus two
// lines of body text all getting pushed apart), not just once around the
// group as a whole.
// `shrinkToFit`/`circular`/`rounded` are stripped too: a shaped badge (the
// only case where a shrinkToFit node itself reaches renderCardShapeDrawing —
// see renderShapedBadge) would otherwise still carry those flags into its own
// txbxContent render, causing renderBlock to promote it to a shape AGAIN
// inside its own shape — infinite recursion, not just a doubled box.
function withoutOwnBoxStyling(node) {
  return { ...node, border: undefined, borderTop: undefined, borderRight: undefined, borderBottom: undefined, borderLeft: undefined, backgroundColor: undefined, backgroundGradient: undefined, padding: undefined, margin: undefined, shrinkToFit: undefined, circular: undefined, rounded: undefined };
}

// Renders one flex/grid layout group as a real `w:tbl` — one column per
// item for flex (a single row), or wrapping into `layoutGroup.columns` per
// row for grid — with each cell carrying its own member(s)' border/
// background as the cell's own w:tcBorders/w:shd rather than each
// paragraph's own (which would only ever span that one text line, not the
// full cell). Equal-width columns split across the page's real content
// width, same "no real layout engine" simplification renderTable's own
// computeColumnWidths already accepts for an unlabeled grid.
//
// The gap between items is real plain, blank spacer columns/rows woven into
// the grid — NOT `w:tblCellSpacing` (WordprocessingML's equivalent of
// HTML's old `cellspacing` attribute), which turned out to render as no gap
// at all in practice. A spacer is just an ordinary borderless, unfilled
// cell/row of the gap's own width/height — every OOXML consumer already
// has to get plain column/row geometry right, so this doesn't depend on a
// less commonly implemented table property the way tblCellSpacing did.
// Matches pptx/slide.js's own BOX_GAP_EMU default (0.1in) — used whenever
// the source CSS declared no `gap` at all, converted to twips (20/pt).
const DEFAULT_LAYOUT_GAP_TWIPS = 144;

// Word applies its own unwritten default table cell margins (left/right,
// typically ~0.08in) whenever a table doesn't declare `w:tblCellMar` — fine
// for a flat-shaded cell (background shading always fills the cell's real
// box regardless of margins), but wrong for a cell holding a real shape
// (renderLayoutRow's needsDrawingShape branch): an inline drawing is laid
// out within the cell's *content* width (colWidth minus that margin), not
// colWidth itself, so a shape sized to the full colWidth silently overflowed
// the cell's right edge — invisible for an inner column (it just overlapped
// the blank gap spacer next to it) but visibly clipped at the page edge for
// the rightmost column. Declaring this explicitly (rather than relying on
// whatever a given Word/LibreOffice build's own unwritten default happens to
// be) lets a shape-holding cell size itself to `colWidth` minus this exact,
// known value instead — see renderLayoutRow.
const LAYOUT_CELL_MARGIN_TWIPS = 108;

function renderLayoutTable(cells, layoutGroup, rctx) {
  const columns = Math.max(1, layoutGroup.mode === 'grid' ? (layoutGroup.columns ?? cells.length) : cells.length);
  const availableWidthTwips = rctx.pageContentWidthTwips ?? CONTENT_WIDTH_TWIPS;
  const gapTwips = layoutGroup.gapPt != null ? Math.round(layoutGroup.gapPt * 20) : DEFAULT_LAYOUT_GAP_TWIPS;
  // CSS `gap` only separates ADJACENT items — (columns - 1) gaps total, not
  // wrapping the row's own outer edges the way tblCellSpacing would have.
  const colWidth = Math.max(1, Math.floor((availableWidthTwips - gapTwips * (columns - 1)) / columns));

  // The real grid alternates item-column, spacer-column, item-column, ... —
  // 2*columns - 1 grid columns in total, one narrow spacer between each
  // pair of real ones (none before the first or after the last, matching
  // CSS `gap`'s own "only between items" behavior).
  const gridWidths = [];
  for (let c = 0; c < columns; c++) {
    if (c > 0) gridWidths.push(gapTwips);
    gridWidths.push(colWidth);
  }

  const b = new XmlBuilder();
  b.el('w:tbl');
  b.el('w:tblPr');
  // No default grid lines — each cell supplies its own border (or none),
  // the same "layout table, not a data table" treatment renderTable's own
  // isLayoutTable branch uses.
  renderTblBorders(b, undefined, ['insideH', 'insideV']);
  b.self('w:tblW', { 'w:w': gridWidths.reduce((a, w) => a + w, 0), 'w:type': 'dxa' });
  b.self('w:tblLayout', { 'w:type': 'fixed' });
  // Made explicit (rather than left to whatever default a given Word/
  // LibreOffice build applies) so renderLayoutRow's shape-holding cells can
  // size themselves against a known value — see LAYOUT_CELL_MARGIN_TWIPS.
  b.el('w:tblCellMar');
  b.self('w:top', { 'w:w': 0, 'w:type': 'dxa' });
  b.self('w:left', { 'w:w': LAYOUT_CELL_MARGIN_TWIPS, 'w:type': 'dxa' });
  b.self('w:bottom', { 'w:w': 0, 'w:type': 'dxa' });
  b.self('w:right', { 'w:w': LAYOUT_CELL_MARGIN_TWIPS, 'w:type': 'dxa' });
  b.up();
  b.up();
  b.el('w:tblGrid');
  for (const w of gridWidths) b.self('w:gridCol', { 'w:w': w });
  b.up();

  const rowCount = Math.ceil(cells.length / columns);
  for (let r = 0; r < rowCount; r++) {
    b.raw(renderLayoutRow(cells.slice(r * columns, r * columns + columns), columns, colWidth, gapTwips, rctx));
    if (r < rowCount - 1) b.raw(renderLayoutSpacerRow(gridWidths, gapTwips));
  }
  b.up(); // w:tbl
  return b.end();
}

// One real row of item cells, with a blank spacer cell of width `gapTwips`
// between each pair — `row[c]` is undefined for a short final grid row's
// missing trailing cells, padded the same empty way as before.
function renderLayoutRow(row, columns, colWidth, gapTwips, rctx) {
  const b = new XmlBuilder();
  b.el('w:tr');
  for (let c = 0; c < columns; c++) {
    if (c > 0) renderPlainCell(b, gapTwips);
    const cellNodes = row[c];
    if (!cellNodes) {
      renderPlainCell(b, colWidth);
      continue;
    }
    const first = cellNodes[0];
    b.el('w:tc');
    // A flex/grid item with its own rounded corner or shadow (a `.box{border-
    // radius:8px}` flat-color tile, a `.card{border-radius:10px;box-shadow:
    // ...}` grid cell) gets the exact same real-shape promotion a standalone
    // card does (see renderCardShapeDrawing/needsDrawingShape) — a plain
    // w:tcBorders/w:shd table cell has no more of a rounded-corner/shadow
    // concept than a plain paragraph border does. The shape's own bodyPr
    // insets replace tcMar's padding, so the cell itself carries no border/
    // shading/margin of its own — just the shape sized to fill it exactly.
    if (needsDrawingShape(first)) {
      b.el('w:tcPr');
      b.self('w:tcW', { 'w:w': colWidth, 'w:type': 'dxa' });
      b.up();
      // The shape must fit the cell's real *content* width — colWidth minus
      // the table's own left+right cell margin (see LAYOUT_CELL_MARGIN_TWIPS)
      // — not the raw column width, or it silently overflows the cell's
      // right edge exactly the way this bug once did.
      const shapeWidthTwips = Math.max(1, colWidth - 2 * LAYOUT_CELL_MARGIN_TWIPS);
      b.raw(renderCardShapeDrawing(cellNodes, rctx, shapeWidthTwips));
    } else {
      const edges = resolveEdgeBorders(first);
      b.el('w:tcPr');
      b.self('w:tcW', { 'w:w': colWidth, 'w:type': 'dxa' });
      // CT_TcPr order: tcBorders, shd, tcMar (same sequence renderTableCellSlot's
      // own tcPr already follows).
      if (hasAnyEdgeBorder(edges)) {
        b.el('w:tcBorders');
        for (const edge of BOX_EDGES) {
          const border = edges[edge];
          if (border) b.self(`w:${edge}`, { 'w:val': border.style, 'w:sz': border.sizeEighthPt, 'w:color': border.color });
        }
        b.up();
      }
      if (first.backgroundColor) b.self('w:shd', { 'w:val': 'clear', 'w:fill': first.backgroundColor });
      renderTcMar(b, first.padding);
      b.up(); // w:tcPr
      b.raw(cellNodes.map((n) => renderBlock(withoutOwnBoxStyling(n), rctx)).join('') || '<w:p/>');
    }
    b.up(); // w:tc
  }
  b.up(); // w:tr
  return b.end();
}

// A borderless, unfilled cell used purely for its own width/height — a
// column spacer between two real cells, or one cell of a spacer row.
function renderPlainCell(b, widthTwips) {
  b.el('w:tc');
  b.el('w:tcPr');
  b.self('w:tcW', { 'w:w': widthTwips, 'w:type': 'dxa' });
  b.up();
  b.raw('<w:p/>');
  b.up();
}

// A full-width blank row of exactly `gapTwips` height, reproducing CSS
// `gap`'s row spacing for a multi-row grid — `w:trHeight`'s `hRule="exact"`
// pins it to that height regardless of the (empty) paragraph Word would
// otherwise pad it out to.
function renderLayoutSpacerRow(gridWidths, gapTwips) {
  const b = new XmlBuilder();
  b.el('w:tr');
  b.el('w:trPr');
  b.self('w:trHeight', { 'w:val': gapTwips, 'w:hRule': 'exact' });
  b.up();
  for (const w of gridWidths) renderPlainCell(b, w);
  b.up(); // w:tr
  return b.end();
}

// ─── Rounded/shadowed boxes as real DrawingML shapes ───────────────────────
// OOXML's plain paragraph border/shading (w:pBdr/w:shd, used everywhere else
// in this file) has no way at all to draw a rounded corner or a drop
// shadow — only needsDrawingShape's two triggers (rounded, shadow) reach
// here; a plain flat-bordered/flat-filled box keeps using the existing
// w:pBdr/w:shd rendering, unchanged. WordprocessingML hosts the exact same
// DrawingML shape vocabulary PPTX uses (`a:prstGeom`/fill/`a:ln`/
// `a:effectLst` — see shared/shape-style.js) via a `<wps:wsp>` element
// (ECMA-376 5th ed. Part 4 §14.2, "wordprocessingShape" — supported by Word
// 2010+, LibreOffice, and Google Docs) inside an ordinary inline drawing,
// the same `w:drawing`/`wp:inline` wrapper already used for images
// (renderImageRun). No `mc:AlternateContent`/VML fallback is emitted for
// pre-2010 Word — this project already accepts the same "modern OOXML
// consumer" assumption for the rest of its DrawingML output (embedded
// images use no VML fallback either).
const NS_WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const EMU_PER_PT = 12700;
const DEFAULT_SHAPE_PADDING_PT = { top: 12, right: 16, bottom: 12, left: 16 };
const SHAPE_BLOCK_SPACING_PT = 6;
// Matches styles.js's own HEADING_SIZES (in half-points: [36,32,28,26,24,22]).
const SHAPE_HEADING_PT = [18, 16, 14, 13, 12, 11];
const SHAPE_DEFAULT_LINE_PT = 14; // ~11pt Normal body text at a 1.2 line-height

// No text-measurement pass exists anywhere in this exporter (matches
// computeContentBasedWidths' own approximation elsewhere) — one line per
// paragraph/heading/quote/code-line, same simplification pptx/shapes.js's
// own estimateCardBlockHeightPt already accepts. Only ever a *starting*
// height: `<a:spAutoFit/>` (see renderCardShapeDrawing) lets Word grow the
// shape to fit its actual rendered text once opened.
function estimateShapeBlockHeightPt(node) {
  if (node.type === 'heading') return (SHAPE_HEADING_PT[node.level - 1] ?? SHAPE_DEFAULT_LINE_PT) * 1.3;
  if (node.type === 'list') return Math.max(node.items?.length ?? 1, 1) * SHAPE_DEFAULT_LINE_PT;
  if (node.type === 'code') return Math.max(node.value.split('\n').length, 1) * SHAPE_DEFAULT_LINE_PT;
  return SHAPE_DEFAULT_LINE_PT * 1.3;
}

function estimateShapeHeightEmu(nodes, padding) {
  const padTop = (padding?.top ?? DEFAULT_SHAPE_PADDING_PT.top) * EMU_PER_PT;
  const padBottom = (padding?.bottom ?? DEFAULT_SHAPE_PADDING_PT.bottom) * EMU_PER_PT;
  const contentHeightPt = nodes.reduce((sum, n) => sum + estimateShapeBlockHeightPt(n) + SHAPE_BLOCK_SPACING_PT, 0);
  return Math.round(contentHeightPt * EMU_PER_PT) + padTop + padBottom;
}

// Renders one cardGroup (or a single self-styled node) as a `<w:p>` holding
// one inline `<wps:wsp>` shape sized to the page's real content width, with
// every member's own paragraphs rendered inside its `<wps:txbx>` text body.
// `nodes[0]` carries the group's border/fill/shadow/rounded (mapElement
// stamps the same values onto every member — see renderCardShape's
// identical "first" convention in pptx/shapes.js).
// `widthTwips` defaults to the full page content width (a standalone card),
// but a flex/grid cell (see renderLayoutRow) passes its own, narrower
// column width instead — the shape then fills that cell exactly like an
// ordinary table-cell box would, just with real rounded corners/shadow/
// gradient. `sizeOverrideEmu` (used only by renderShapedBadge) skips
// both size computations entirely — a shrink-to-fit circle/oval/pill badge
// needs its own small, content-fit size, never the page/column width a card
// stretches to. `bodyAnchor`/`bodyWrap` default to a card's own convention
// (top-anchored, word-wrapping prose) — renderShapedBadge overrides
// both to match a badge's centered, single-line text instead.
function renderCardShapeDrawing(nodes, rctx, widthTwips, { sizeOverrideEmu, bodyAnchor = 't', bodyWrap = 'square' } = {}) {
  const first = nodes[0];
  const padding = first.padding ?? DEFAULT_SHAPE_PADDING_PT;
  const cx = sizeOverrideEmu?.cx ?? (widthTwips ?? rctx.pageContentWidthTwips ?? CONTENT_WIDTH_TWIPS) * 635; // twips → EMU
  // An explicit CSS `height` on the source container (e.g. a full-height
  // `.hero{height:900px}` banner) is a real authored size — same "explicit
  // beats inferred" precedence renderShapedBadge's own widthPx/heightPx
  // handling already uses for a badge — so it wins over the rough per-block
  // estimate below (never shrinking under it, so real wrapped text is never
  // clipped by an under-estimated explicit height).
  const explicitCy = first.heightPx != null ? Math.round(first.heightPx * EMU_PER_PX) : undefined;
  const estimatedCy = estimateShapeHeightEmu(nodes, padding);
  const cy = sizeOverrideEmu?.cy ?? (explicitCy != null ? Math.max(explicitCy, estimatedCy) : estimatedCy);
  const docPrId = rctx.nextImageId();

  const b = new XmlBuilder();
  b.el('w:p');
  // The group's own outer margin (see withoutOwnBoxStyling's own note on why
  // it's stripped from each member) belongs here instead — spacing around
  // the shape as a whole, exactly like any other paragraph's w:spacing.
  if (first.margin) {
    b.el('w:pPr');
    renderSpacing(b, first.margin);
    b.up();
  }
  b.el('w:r');
  b.el('w:drawing');
  b.el('wp:inline', { distT: 0, distB: 0, distL: 0, distR: 0 });
  b.self('wp:extent', { cx, cy });
  b.self('wp:effectExtent', { l: 0, t: 0, r: 0, b: 0 });
  b.self('wp:docPr', { id: docPrId, name: `Card ${docPrId}` });
  b.self('wp:cNvGraphicFramePr');
  b.el('a:graphic', { 'xmlns:a': NS.a });
  b.el('a:graphicData', { uri: NS_WPS });
  b.el('wps:wsp', { 'xmlns:wps': NS_WPS });
  b.self('wps:cNvSpPr');
  b.el('wps:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x: 0, y: 0 });
  b.self('a:ext', { cx, cy });
  b.up();
  // A real CSS border-radius earns a rounded shape even with no `border` at
  // all (a borderless gradient panel with its own border-radius), and
  // border-radius:50%+ earns a real ellipse over either — matches
  // pptx/shapes.js's identical renderCardShape decision (see
  // cardPresetGeometry).
  b.el('a:prstGeom', { prst: cardPresetGeometry(first) });
  b.self('a:avLst');
  b.up();
  b.raw(shapeFillXml(first));
  b.raw(shapeLineXml(first.border));
  if (first.shadow) b.raw(shadowEffectXml());
  b.up(); // wps:spPr
  b.el('wps:txbx');
  b.el('w:txbxContent');
  // Each member's own border/background/padding is already the shape's own
  // fill/line above — rendering its paragraph with them still set would
  // additionally draw that same box a second time, as a w:pBdr/w:shd nested
  // inside the shape's text body.
  const contentNodes = nodes.map((n) => withoutOwnBoxStyling(n));
  // A badge (sizeOverrideEmu set) has no w:spacing of its own after
  // withoutOwnBoxStyling strips margin to undefined — renderSpacing then
  // emits nothing at all, so the paragraph silently inherits the Normal
  // style's own default `w:after="160"` (8pt — see styles.js). That extra,
  // invisible trailing space becomes part of what `anchor="ctr"` centers,
  // and since it only ever sits *after* the text, centering "text + 8pt of
  // blank space" pushes the visible text itself up from true center —
  // exactly the badge-text-sits-near-the-top symptom this was chasing.
  // Explicitly zeroing it here (only for a badge's own single paragraph, not
  // a general multi-block card, where that breathing room between a heading
  // and its body text is expected) removes the invisible offset entirely.
  if (sizeOverrideEmu) {
    for (const n of contentNodes) n.margin = { beforePt: 0, afterPt: 0 };
  }
  b.raw(contentNodes.map((n) => renderBlock(n, rctx)).join(''));
  b.up(); // w:txbxContent
  b.up(); // wps:txbx
  b.el('wps:bodyPr', {
    wrap: bodyWrap, anchor: bodyAnchor,
    lIns: Math.round((padding.left ?? DEFAULT_SHAPE_PADDING_PT.left) * EMU_PER_PT),
    tIns: Math.round((padding.top ?? DEFAULT_SHAPE_PADDING_PT.top) * EMU_PER_PT),
    rIns: Math.round((padding.right ?? DEFAULT_SHAPE_PADDING_PT.right) * EMU_PER_PT),
    bIns: Math.round((padding.bottom ?? DEFAULT_SHAPE_PADDING_PT.bottom) * EMU_PER_PT),
  });
  // A badge (sizeOverrideEmu set — see renderShapedBadge) already has an
  // exact, correct size computed from real font metrics, so it never needs
  // Word to resize it — and `a:spAutoFit` paired with `anchor="ctr"` on a
  // wps:wsp shape has a real, observed Word-specific bug: opening the file
  // triggers Word's own autofit re-layout, which — unlike PowerPoint's
  // native p:sp shapes — grows/settles the shape without honoring the
  // vertical-center anchor afterward, leaving single-line badge text sitting
  // near the top with visible empty space below instead of centered. A
  // plain card's height, by contrast, is only a rough estimate (see
  // estimateShapeHeightEmu's own module doc), so it still needs
  // `a:spAutoFit` to grow to whatever the real text turns out to need — an
  // explicit CSS height (explicitCy) is the same "already exact, don't let
  // Word resize it" case as a badge: `a:spAutoFit` would otherwise silently
  // settle the shape back down to content height, the same way it would
  // discard a badge's own computed size.
  b.self(sizeOverrideEmu || explicitCy != null ? 'a:noAutofit' : 'a:spAutoFit');
  b.up(); // wps:bodyPr
  return b.end(); // auto-closes wps:wsp, a:graphicData, a:graphic, wp:inline, w:drawing, w:r, w:p
}

// ─── Circular/oval shrink-to-fit badges as real DrawingML shapes ──────────
// renderShrinkToFitBox's 1-cell autofit table is the right tool for an
// ordinary rounded pill (a table cell is a fine visual stand-in for
// "roughly rounded box" even though it can't actually curve its corners) —
// but a table cell is *always* rectangular, so it can never represent a
// genuine circle/oval (border-radius:50%+) at all; that case needs the same
// real `ellipse` DrawingML shape PPTX already draws for it (see
// pptx/shapes.js's renderStyledBoxShape). Matches Word's own default body
// font (11pt Calibri) when the badge declares no explicit size/family of its
// own — same fallback convention pptx/shapes.js's resolveOwnFontSizePt uses.
const DEFAULT_BADGE_FONT_PT = 11;
const DEFAULT_BADGE_PADDING_PT = { top: 6, right: 10, bottom: 6, left: 10 };

function resolveBadgeFontSizePt(node, fallback = DEFAULT_BADGE_FONT_PT) {
  const only = node.children?.length === 1 ? node.children[0] : undefined;
  return only?.type === 'styledSpan' && only.style?.fontSizePt ? only.style.fontSizePt : fallback;
}

function resolveBadgeFontFamily(node, fallback = 'Calibri') {
  const only = node.children?.length === 1 ? node.children[0] : undefined;
  return only?.type === 'styledSpan' && only.style?.fontFamily ? resolveFontFamilyName(only.style.fontFamily) : fallback;
}

// Sized the same way pptx/shapes.js's renderStyledBoxShape sizes a badge:
// an explicit CSS width/height (the common case for a real circle — a
// circle with no declared size has no defined diameter) wins outright;
// otherwise font-metrics.js's single-line auto-size estimate (with its own
// built-in clipping-safety bias) provides a reasonable fallback for an
// unsized badge (e.g. a small inline status dot). Handles any shrink-to-fit
// badge with a shape-worthy corner style, not just a true circle/oval:
// `cardPresetGeometry` (inside renderCardShapeDrawing below) already picks
// 'ellipse' for `circular` or 'roundRect' for a merely-`rounded` badge (e.g.
// `.badge{border-radius:20px}`, a pill) on its own — a shrink-to-fit badge
// with *no* rounding at all still renders as a plain square-cornered table
// via renderShrinkToFitBox instead, since a real DrawingML shape buys
// nothing over that for a square box.
function renderShapedBadge(node, rctx) {
  const text = flattenInline(node.children ?? []);
  const fontSizePt = resolveBadgeFontSizePt(node);
  const fontFamily = resolveBadgeFontFamily(node);
  const paddingPt = node.padding ?? DEFAULT_BADGE_PADDING_PT;
  const estimated = estimateShapeSize(text, fontSizePt, fontFamily, { padding: paddingPt });

  const cx = node.widthPx != null ? Math.round(node.widthPx * EMU_PER_PX) : Math.round(estimated.width * EMU_PER_PT);
  const cy = node.heightPx != null ? Math.round(node.heightPx * EMU_PER_PX) : Math.round(estimated.height * EMU_PER_PT);

  // wrap:'none' (PPTX's own convention for a badge — see
  // pptx/shapes.js's renderStyledBoxShape) is NOT used here: PowerPoint's
  // native p:sp shapes honor it correctly, but Word's wps:wsp — a less
  // mature, WordprocessingML-hosted DrawingML shape — does not reliably
  // honor wrap:'none' paired with an explicit fixed size in real-world
  // renderers, observed breaking text into one character per line. 'square'
  // is the same standard word-wrap mode the card shapes already use
  // successfully, and a properly-sized box never needs to actually wrap a
  // short badge label.
  // A circular badge's label defaults to centered horizontally too
  // (bodyAnchor:'ctr' above only centers vertically) — the source CSS
  // rarely declares an explicit text-align for a circle (real examples
  // center it via flexbox, e.g. `.circle{display:flex;align-items:center;
  // justify-content:center}`, which carries no equivalent onto the plain
  // text child at all), and a short label left-aligned inside a much wider
  // circle reads as visibly broken, not just imprecise. An explicit CSS
  // text-align still wins.
  const centered = { ...node, align: node.align ?? 'center' };
  return renderCardShapeDrawing([centered], rctx, undefined, { sizeOverrideEmu: { cx, cy }, bodyAnchor: 'ctr', bodyWrap: 'square' });
}

function renderBlock(node, rctx) {
  switch (node.type) {
    case 'paragraph':
      // A merely-`rounded` badge (e.g. `.badge{border-radius:20px}`, a pill —
      // not necessarily `circular`, which only covers 50%+) needs the same
      // real DrawingML shape a circle does: renderShrinkToFitBox's plain
      // table has no OOXML concept of rounded table-cell corners at all, so
      // routing it there instead would silently square off every non-circle
      // rounded badge.
      if (node.shrinkToFit && (node.circular || node.rounded)) return renderShapedBadge(node, rctx);
      return node.shrinkToFit ? renderShrinkToFitBox(node, rctx) : renderParagraph(node, rctx);
    case 'heading':
      return renderHeading(node, rctx);
    case 'thematicBreak':
      return renderThematicBreak();
    case 'pageBreak':
      return renderPageBreak();
    case 'quote':
      return renderQuote(node, rctx);
    case 'code':
      return renderCode(node);
    case 'list':
      return renderList(node, rctx, 0);
    case 'table':
      return renderTable(node, rctx);
    case 'image':
      return renderImageParagraph(node, rctx);
    case 'figure':
      return renderFigure(node, rctx);
    case 'svgShapes':
      return renderSvgShapesParagraph(node, rctx);
    // <section>/<article>/<main> (see dom-to-idm.js's mapElementCore) wrap
    // arbitrary block content — recursing through the same renderBlocks used
    // for the document root (and every quote's own children) is what
    // actually handles that heterogeneity (a section can itself contain
    // headings, lists, tables, nested cardGroups...). This case was missing
    // entirely until now, which silently dropped every <section>'s content:
    // renderBlock's own `default: return ''` swallowed the whole node with
    // no warning. A section's own border/background/padding (now correctly
    // attached — see dom-to-idm.js's BOX_STYLE_NODE_TYPES) still isn't
    // rendered as a wrapping box here; unlike quote's single paragraph
    // child, a section's children can be lists/tables/nested quotes that
    // don't have a paragraph-level w:pBdr/w:shd to hang it on, so that part
    // remains a known gap, not silently-dropped content.
    case 'section':
      return renderBlocks(node.children, rctx);
    default:
      return '';
  }
}

// A <div class="box">'s own border(s) — either a uniform one (CSS `border`,
// applies equally to all 4 sides) or a per-side one (`border-left:5px solid
// #hex`, the common "accent stripe" pattern: a colored bar down one edge
// with no border on the other three — see dom-to-idm.js's mapElement for
// both). Resolved independently per edge, in priority order: a genuinely
// side-specific declaration > the uniform `border` > a synthesized invisible
// border matching the box's own background fill, used purely to unlock
// w:pBdr's per-edge padding space (`w:space` — one of the few OOXML
// paragraph properties with a direct "padding inside a border" concept) when
// there's a background+padding but no real border declared on that
// particular edge — the same fake-border trick used to pad a shaded box by
// hand in raw OOXML. An edge with no border at all (real or synthesized)
// renders nothing, matching CSS's "no border declared on this side" exactly.
function resolveEdgeBorders(node) {
  const invisible = node.backgroundColor && node.padding
    ? { color: node.backgroundColor, sizeEighthPt: 2, style: 'single' }
    : undefined;
  return {
    top: node.borderTop ?? node.border ?? invisible,
    right: node.borderRight ?? node.border ?? invisible,
    bottom: node.borderBottom ?? node.border ?? invisible,
    left: node.borderLeft ?? node.border ?? invisible,
  };
}

function hasAnyEdgeBorder(edges) {
  return Boolean(edges.top || edges.right || edges.bottom || edges.left);
}

// `w:space` (distance from the border line to the text) comes from the
// element's own CSS padding per side when set, falling back to a small fixed
// default otherwise. Real ST_BorderSpace is 0-31pt, so it's clamped to that
// range. An edge with no border in `edges` is omitted entirely.
function renderPBdr(b, edges, padding, hasBorder) {
  if (!hasBorder) return;
  b.el('w:pBdr');
  for (const edge of BOX_EDGES) {
    const border = edges[edge];
    if (!border) continue;
    const spacePt = padding?.[edge];
    const space = spacePt !== undefined ? Math.max(0, Math.min(31, Math.round(spacePt))) : 1;
    b.self(`w:${edge}`, { 'w:val': border.style, 'w:sz': border.sizeEighthPt, 'w:space': space, 'w:color': border.color });
  }
  b.up();
}

// A padded box with no border on any edge (real or synthesized) has no
// OOXML "padding" construct to hang space off of at all — w:ind (indent)
// approximates just the horizontal component by insetting the whole
// paragraph from the margins. Skipped whenever any edge has a border:
// indenting would shift that border line itself inward rather than adding
// space between it and the text, which is what renderPBdr's w:space already
// does correctly.
function renderIndent(b, node, hasBorder) {
  if (hasBorder || !node.padding) return;
  const left = node.padding.left !== undefined ? Math.round(node.padding.left * 20) : undefined;
  const right = node.padding.right !== undefined ? Math.round(node.padding.right * 20) : undefined;
  if (left === undefined && right === undefined) return;
  const attrs = {};
  if (left !== undefined) attrs['w:left'] = left;
  if (right !== undefined) attrs['w:right'] = right;
  b.self('w:ind', attrs);
}

// <div style="margin-top/-bottom">'s own vertical margin — direct formatting
// here overrides whatever spacing the paragraph's style (Normal/HeadingN)
// would otherwise apply, the same override relationship a browser's author
// CSS has over its UA-stylesheet default margin. Only the sides actually
// declared are emitted; an omitted side still inherits the style's spacing.
function renderSpacing(b, margin) {
  if (!margin) return;
  const attrs = {};
  if (margin.beforePt !== undefined) attrs['w:before'] = Math.round(margin.beforePt * 20); // pt → twips
  if (margin.afterPt !== undefined) attrs['w:after'] = Math.round(margin.afterPt * 20);
  b.self('w:spacing', attrs);
}

// The border/shading/spacing/indent/alignment sequence shared by every plain
// paragraph and every heading's own w:pPr — same fields, same order, same
// resolveEdgeBorders/hasAnyEdgeBorder inputs either way; a heading's w:pPr
// additionally opens with w:pStyle (see renderHeading), which isn't part of
// this shared tail.
function renderBoxStylingPPr(b, node, edges, hasBorder) {
  renderPBdr(b, edges, node.padding, hasBorder);
  if (node.backgroundColor) b.self('w:shd', { 'w:val': 'clear', 'w:fill': node.backgroundColor });
  renderSpacing(b, node.margin);
  renderIndent(b, node, hasBorder);
  if (node.align) b.self('w:jc', { 'w:val': node.align });
}

function renderParagraph(node, rctx) {
  const b = new XmlBuilder();
  b.el('w:p');
  const hasBoxStyling = node.align || node.border || node.borderTop || node.borderRight ||
    node.borderBottom || node.borderLeft || node.backgroundColor || node.margin || node.padding;
  if (hasBoxStyling) {
    const edges = resolveEdgeBorders(node);
    b.el('w:pPr');
    renderBoxStylingPPr(b, node, edges, hasAnyEdgeBorder(edges));
    b.up();
  }
  b.raw(renderRuns(node.children, {}, rctx));
  return renderBookmarked(node, b.end(), rctx);
}

// Builds a `w:tblBorders` block for a table's outer perimeter — either a
// real border (a data table's `table.border` override, or a shrink-to-fit
// button box's own border) applied to all 4 edges, or `'none'` on all 4 when
// omitted (a borderless layout table, or a borderless button box). `insideEdges`
// (when passed) additionally forces those inside-line edges to 'none' too —
// an outer border on a <table> never implies a grid between cells (see
// renderTable's own comment on this), and a bare layout table has no visible
// lines at all; a shrink-to-fit box has no inside edges to speak of (a single
// cell), so it omits this parameter entirely.
function renderTblBorders(b, border, insideEdges) {
  b.el('w:tblBorders');
  for (const edge of BOX_EDGES) {
    if (border) b.self(`w:${edge}`, { 'w:val': border.style, 'w:sz': border.sizeEighthPt, 'w:color': border.color });
    else b.self(`w:${edge}`, { 'w:val': 'none' });
  }
  for (const edge of insideEdges ?? []) b.self(`w:${edge}`, { 'w:val': 'none' });
  b.up();
}

// A "shrink to fit" paragraph (currently only <button> — see dom-to-idm.js's
// mapElementCore) renders as a single-cell, autofit table instead of a plain
// shaded paragraph. OOXML paragraph borders/shading (renderPBdr/w:shd) always
// span the full text column width — there's no "size to content" option for
// a paragraph, unlike a CSS inline-block button that's naturally sized to
// its own text+padding. A borderless table with `w:tblW type="auto"` +
// `w:tblLayout type="autofit"` DOES size to its content, and the whole table
// can be centered on the page via `w:jc` — the standard technique other DOCX
// generators use for an auto-sized colored badge/button.
function renderShrinkToFitBox(node, rctx) {
  const b = new XmlBuilder();
  b.el('w:tbl');
  b.el('w:tblPr');
  b.self('w:tblW', { 'w:w': 0, 'w:type': 'auto' });
  b.self('w:tblLayout', { 'w:type': 'autofit' });
  if (node.align) b.self('w:jc', { 'w:val': node.align });
  renderTblBorders(b, node.border, ['insideH', 'insideV']);
  b.up(); // w:tblPr
  // CT_Tbl requires exactly one w:tblGrid right after w:tblPr — every other
  // table this exporter builds (renderTable/renderLayoutTable) already
  // includes one, but this one didn't, making the table schema-invalid.
  // A missing/invalid grid is exactly the condition real Word's own
  // "AutoFit to Contents" behavior doesn't handle gracefully: with no column
  // width to reason about, Word falls back to stretching the table across
  // the full page width instead of sizing to content — the badge/button
  // rendering as a full-width bar instead of a small pill. A single bare
  // `w:gridCol` (no `w:w`) is schema-valid and leaves the actual width to
  // `w:tblLayout type="autofit"` to determine from content, same as intended.
  b.el('w:tblGrid');
  b.self('w:gridCol');
  b.up();

  b.el('w:tr');
  b.el('w:tc');
  b.el('w:tcPr');
  b.self('w:tcW', { 'w:w': 0, 'w:type': 'auto' });
  // CT_TcPr is a strict, order-sensitive schema — shd must precede tcMar (see
  // renderTableCellSlot's own tcPr, which follows the same tcBorders → shd →
  // tcMar → textDirection order).
  if (node.backgroundColor) b.self('w:shd', { 'w:val': 'clear', 'w:fill': node.backgroundColor });
  renderTcMar(b, node.padding);
  b.up(); // w:tcPr
  b.el('w:p');
  b.el('w:pPr');
  // Every paragraph in the document inherits the Normal style's own default
  // "space after" (see styles.js — 160 twips, meant to give ordinary body
  // paragraphs breathing room). Left alone, that adds on top of the button's
  // own tcMar padding, making the box visibly taller than its CSS padding
  // actually specifies — the box's size should come from that padding alone.
  b.self('w:spacing', { 'w:before': 0, 'w:after': 0 });
  if (node.align) b.self('w:jc', { 'w:val': node.align });
  b.up();
  b.raw(renderRuns(node.children, {}, rctx));
  b.up(); // w:p
  b.up(); // w:tc
  b.up(); // w:tr
  b.up(); // w:tbl

  return renderBookmarked(node, b.end(), rctx);
}

function renderHeading(node, rctx) {
  const b = new XmlBuilder();
  const edges = resolveEdgeBorders(node);
  b.el('w:p');
  b.el('w:pPr');
  b.self('w:pStyle', { 'w:val': `Heading${node.level}` });
  renderBoxStylingPPr(b, node, edges, hasAnyEdgeBorder(edges));
  b.up();
  b.raw(renderRuns(node.children, {}, rctx));
  return renderBookmarked(node, b.end(), rctx);
}

// A bottom paragraph border is the standard way to render a horizontal rule
// in WordprocessingML — there's no dedicated "hr" construct.
function renderThematicBreak() {
  const b = new XmlBuilder();
  b.el('w:p');
  b.el('w:pPr');
  b.el('w:pBdr');
  b.self('w:bottom', { 'w:val': 'single', 'w:sz': 6, 'w:space': 1, 'w:color': 'auto' });
  b.up();
  b.up();
  return b.end();
}

function renderPageBreak() {
  const b = new XmlBuilder();
  b.el('w:p');
  b.el('w:r');
  b.self('w:br', { 'w:type': 'page' });
  b.up();
  return b.end();
}

// Each paragraph inside the quote gets the real Quote style applied directly
// — no separate empty stub paragraph (the exact bug found in the prior
// attempt). Non-paragraph children (a nested list/table/quote) render as
// themselves; forcing an arbitrary block into a paragraph style makes no sense.
// The blockquote's own border/backgroundColor/padding (e.g. `blockquote{
// border-left:6px solid #2563EB;padding:15px;background:#eef4ff}` — see
// dom-to-idm.js's mapElement quote branch) live on `node` itself, not on its
// paragraph children, so they're passed down into renderQuoteChild/
// renderParagraphStyled rather than read off the child being rendered.
function renderQuote(node, rctx) {
  return node.children.map((child) => renderQuoteChild(child, node, rctx)).join('');
}

function renderQuoteChild(node, quoteNode, rctx) {
  if (node.type === 'paragraph') return renderParagraphStyled(node, 'Quote', quoteNode, rctx);
  return renderBlock(node, rctx);
}

function renderParagraphStyled(node, styleId, quoteNode, rctx) {
  const b = new XmlBuilder();
  // Merges the quote's own box styling with this specific child's own
  // align — reuses the exact same resolveEdgeBorders/renderBoxStylingPPr
  // helpers a plain bordered/shaded div paragraph already uses (see
  // renderParagraph), so a synthesized invisible border still unlocks
  // w:pBdr's per-edge `w:space` for the padding even when the source only
  // declared `border-left` (the common blockquote accent-stripe pattern).
  const styled = { ...quoteNode, align: node.align };
  const edges = resolveEdgeBorders(styled);
  b.el('w:p');
  b.el('w:pPr');
  b.self('w:pStyle', { 'w:val': styleId });
  renderBoxStylingPPr(b, styled, edges, hasAnyEdgeBorder(edges));
  b.up();
  b.raw(renderRuns(node.children, {}, rctx));
  return b.end();
}

// One paragraph, shaded, monospace; multi-line code joins lines with w:br
// rather than one <w:p> per line, so it stays visually one block. A <pre>
// with its own explicit background/color (a dark-theme code block, e.g.
// `pre{background:#272822;color:#eee}`) uses that instead of the plain
// light-gray default — passed through to the runs too (not just the
// paragraph shading) so textRun's own `code`-default fill doesn't compete
// with it (see renderRPr: `code && !backgroundColor` is the only case that
// applies the F1F1F1 fallback).
function renderCode(node) {
  const b = new XmlBuilder();
  const edges = resolveEdgeBorders(node);
  const hasBorder = hasAnyEdgeBorder(edges);
  b.el('w:p');
  b.el('w:pPr');
  // A wrapping `<div style="border:...;padding:...">` around the `<pre>`
  // (see dom-to-idm.js's BOX_STYLE_NODE_TYPES) now attaches border/padding
  // to this same node — rendered the same w:pBdr/w:ind way any other
  // bordered/padded box is. Shading still defaults to backgroundColor
  // (falling back to the plain gray F1F1F1) regardless of border, unlike a
  // plain paragraph's box styling, since a code block always wants *some*
  // visible fill even with no CSS background at all.
  renderPBdr(b, edges, node.padding, hasBorder);
  b.self('w:shd', { 'w:val': 'clear', 'w:fill': node.backgroundColor ?? 'F1F1F1' });
  renderIndent(b, node, hasBorder);
  b.up();
  const rProps = { code: true, backgroundColor: node.backgroundColor, color: node.color };
  const lines = node.value.split('\n');
  lines.forEach((line, i) => {
    if (i > 0) {
      b.el('w:r');
      b.self('w:br');
      b.up();
    }
    b.raw(textRun(line, rProps));
  });
  return b.end();
}

// ─── Lists (W5) ─────────────────────────────────────────────────────────────
// Every distinct <ul>/<ol> — including a nested one — registers its own numId
// via NumberingManager, so Word treats separate lists as independently
// numbered. Nesting depth becomes w:ilvl, which (since each numId's
// abstractNum defines all 9 levels) also compounds indent and, for bullets,
// cycles the marker character deeper.

function renderList(node, rctx, depth) {
  const numId = rctx.numbering.registerList(node.ordered, node.start);
  const styleId = node.ordered ? 'ListNumber' : 'ListBullet';
  // This specific list node's own border/background/padding, if it has any
  // (see dom-to-idm.js's BOX_STYLE_NODE_TYPES) — applied to every one of
  // *its own* items' paragraphs, relying on Word's own border-merging
  // between adjacent identically-bordered paragraphs to still read as one
  // box, the same accepted approximation renderBlocks' own flat-bordered
  // cardGroup fallback already uses. Passed unconditionally, not gated on
  // depth: whether a `<ul>`/`<ol>` picked up styling has nothing to do with
  // how deeply nested it is, only with whether *this* list (top-level or a
  // `<ul style="...">` nested several levels down) is the one a wrapping div
  // — or its own inline style — actually styled. A nested sub-list that
  // wasn't itself styled simply has border/backgroundColor/padding all
  // undefined, so this is a no-op for the common case.
  return node.items.map((item) => renderListItem(item, rctx, depth, numId, styleId, node)).join('');
}

function renderListItem(item, rctx, depth, numId, styleId, boxStyle) {
  const [first, ...rest] = item.children;
  let out = '';

  if (first?.type === 'paragraph') {
    out += renderNumberedParagraph(first, depth, numId, styleId, rctx, boxStyle);
  } else if (first) {
    out += renderBlock(first, rctx);
  }

  for (const child of rest) {
    out += child.type === 'list'
      ? renderList(child, rctx, depth + 1)
      : renderBlock(child, rctx);
  }

  return out;
}

function renderNumberedParagraph(node, depth, numId, styleId, rctx, boxStyle) {
  const b = new XmlBuilder();
  const edges = boxStyle ? resolveEdgeBorders(boxStyle) : undefined;
  const hasBorder = Boolean(edges && hasAnyEdgeBorder(edges));
  b.el('w:p');
  b.el('w:pPr');
  b.self('w:pStyle', { 'w:val': styleId });
  b.el('w:numPr');
  b.self('w:ilvl', { 'w:val': depth });
  b.self('w:numId', { 'w:val': numId });
  b.up();
  // CT_PPr requires numPr before pBdr/shd/ind — see renderTableCellSlot's
  // own note on CT_TcPr's equally strict ordering.
  if (boxStyle) {
    renderPBdr(b, edges, boxStyle.padding, hasBorder);
    if (boxStyle.backgroundColor) b.self('w:shd', { 'w:val': 'clear', 'w:fill': boxStyle.backgroundColor });
    renderIndent(b, boxStyle, hasBorder);
  }
  if (node.align) b.self('w:jc', { 'w:val': node.align });
  b.up();
  b.raw(renderRuns(node.children, {}, rctx));
  return b.end();
}

// ─── Tables (W6) ────────────────────────────────────────────────────────────
// The IDM's TableRowNode.cells — like HTML — omits a cell in a later row once
// a prior row's rowspan covers it. OOXML requires every <w:tr> to have the
// same number of <w:tc>/gridSpan-widths as <w:tblGrid> has columns, so the
// collapsed grid is expanded back out: a rowspan's covered rows get a real
// placeholder <w:tc> with a bare <w:vMerge/> (continuation) matching the
// origin cell's gridSpan, and the origin cell gets <w:vMerge w:val="restart"/>.
// This is exactly the gap that corrupted tables in the prior DOCX attempt.

const CONTENT_WIDTH_TWIPS = 9360; // ~6.5in, matches the W1 page margins

// availableWidthTwips is the full page content width at the top level, or
// (for a table nested inside a cell) that cell's own allotted width — without
// this, a nested table always claimed the full page width for itself,
// forcing Word to blow out its containing cell/column just like the
// unscaled image bug did one level up.
// Defaults to rctx.pageContentWidthTwips (the real page content width for
// this export, set once in buildDocumentXml from convert()'s pageSize/
// orientation/margins options) — CONTENT_WIDTH_TWIPS is only a fallback for
// callers that never went through buildDocumentXml at all (e.g. a test
// constructing a bare rctx by hand).
function renderTable(table, rctx, availableWidthTwips = rctx.pageContentWidthTwips ?? CONTENT_WIDTH_TWIPS) {
  const headRows = table.head ?? [];
  const allRows = [...headRows, ...table.rows];
  if (allRows.length === 0) return '';

  const { grid, columnCount } = expandTableGrid(allRows);
  // `table.widthPercent` (from CSS `width:%` on the <table> itself, e.g.
  // `width:50%`) narrows the space columns are computed against — without
  // this, computeContentBasedWidths' own leftover-space distribution always
  // stretches a naturally-narrow table (like a 2-column totals block) to
  // fill the entire page width, which is what made it render full-width and
  // flush-left regardless of the source HTML's intended half-width table.
  const tableWidthTwips = table.widthPercent != null
    ? Math.max(1, Math.round((table.widthPercent / 100) * availableWidthTwips))
    : availableWidthTwips;
  const colWidths = computeColumnWidths(grid, columnCount, tableWidthTwips);

  // A table whose cells hold block content (paragraphs, a nested table), or
  // whose cells are all empty (a decorative spacer/divider row), is being
  // used for layout/grouping rather than tabular data (the same signal
  // TXT/Markdown use to pick flowing text over an ASCII/GFM grid) — showing
  // visible borders on that turns a logo/metadata/card layout into a page of
  // empty boxes in Word. Only a real data table — genuinely non-empty, plain
  // inline cell values — gets the bordered TableGrid look. An explicit
  // `border` declared on the source <table> always wins over this heuristic
  // — the same "explicit CSS beats inference" rule used for run styling.
  // That default is further narrowed to a table with NO styling anywhere
  // (tableHasExplicitStyling) — the common case this heuristic targets is
  // content pasted with zero CSS at all, where Word's default grid is a
  // sensible guess; a table the author DID style in some way (alignment,
  // per-cell borders/padding/width, ...) but deliberately left borderless
  // (e.g. a plain totals block under a real bordered data table) has that
  // choice respected instead of silently overridden.
  const skipDefaultBorder = !isDataTable(allRows) || tableHasExplicitStyling(table, allRows);

  // CT_TblPrBase's child sequence is fixed by the schema: w:tblStyle, then
  // (much later) w:tblW, w:jc, w:tblCellSpacing, w:tblInd, w:tblBorders,
  // w:shd, w:tblLayout, ... — Word validates this strictly enough that an
  // element emitted out of order (tblBorders used to be written before
  // tblW/jc here) can silently fail to apply, which is what made w:jc go
  // ignored even though the value itself was correct.
  const b = new XmlBuilder();
  b.el('w:tbl');
  b.el('w:tblPr');
  if (!table.border && !skipDefaultBorder) {
    b.self('w:tblStyle', { 'w:val': 'TableGrid' });
  }
  b.self('w:tblW', { 'w:w': colWidths.reduce((a, w) => a + w, 0), 'w:type': 'dxa' });
  // `margin-left:auto`/`margin-right:auto` (dom-to-idm.js's mapTable) is the
  // standard way a narrower-than-page table is pushed to one side or
  // centered without flex/grid — Word's own `w:jc` (table justification)
  // does exactly that for a `w:tbl` whose own width is less than the page's.
  if (table.marginAuto === 'left') b.self('w:jc', { 'w:val': 'right' }); // hugs the right edge
  else if (table.marginAuto === 'both') b.self('w:jc', { 'w:val': 'center' });
  if (table.border) {
    // CSS `border` on a <table> element styles only the table's own outer
    // perimeter — it does not add lines between rows/cells (that would need
    // border-collapse plus a border on each <td>, which isn't modeled here).
    // Applying it to insideH/insideV too drew a spurious line between, e.g.,
    // a card's colored top strip and its content row that the source never had.
    renderTblBorders(b, table.border, ['insideH', 'insideV']);
  } else if (skipDefaultBorder) {
    renderTblBorders(b, undefined, ['insideH', 'insideV']);
  }
  // Fixed layout: Word must respect the widths below rather than re-flowing
  // columns to fit an oversized image or long text — the exact bug that made
  // a wide image blow out its column and crush the others.
  b.self('w:tblLayout', { 'w:type': 'fixed' });
  b.up();

  b.el('w:tblGrid');
  for (const w of colWidths) b.self('w:gridCol', { 'w:w': w });
  b.up();

  grid.forEach((gridRow, rowIndex) => {
    b.raw(renderTableRow(gridRow, rowIndex < headRows.length, colWidths, rctx));
  });

  b.up();
  return b.end();
}

// A layout table's first row typically carries each column's real width as
// a percentage (<td width="52%">, common in email/CMS table layouts) — when
// every cell in row 0 declares one, distribute the page's content width by
// those proportions instead of dividing it equally, which is what let a
// wide image column crush its neighbors down to illegibly narrow columns.
// Falls back to computeContentBasedWidths (not a flat equal split) when no
// column declares a width at all — see that function's own comment.
function computeColumnWidths(grid, columnCount, availableWidthTwips) {
  const firstRow = grid[0] ?? [];
  const allSpecified = firstRow.length > 0 && firstRow.every((slot) => slot.cell?.widthPercent != null);
  const totalPercent = firstRow.reduce((sum, slot) => sum + (slot.cell?.widthPercent ?? 0), 0);

  if (!allSpecified || totalPercent <= 0) {
    return computeContentBasedWidths(grid, columnCount, availableWidthTwips);
  }

  const widths = [];
  for (const slot of firstRow) {
    const totalTwips = Math.round((slot.cell.widthPercent / totalPercent) * availableWidthTwips);
    const perCol = Math.max(1, Math.floor(totalTwips / slot.colspan));
    for (let i = 0; i < slot.colspan; i++) widths.push(perCol);
  }
  while (widths.length < columnCount) widths.push(Math.floor(availableWidthTwips / columnCount));
  return widths.slice(0, columnCount);
}

// A cell's own font-size (e.g. `<td style="font-size:60px">INVOICE</td>`, a
// large title banner) — character count alone badly underestimates how much
// horizontal room such a cell actually needs, especially once colspan
// divides its already-short text across several columns. Only checks the
// cell's top-level children for a single styledSpan carrying fontSizePt —
// the shape produced when the whole cell's own inline style wraps all of
// its content (mapInlineWithStyle) — not a deep/exhaustive search across
// mixed-formatting content, since that's the overwhelmingly common real
// case for a cell that's specifically sized as a banner/heading.
function cellFontScale(cell) {
  const nodes = cell.children ?? [];
  const fontSizePt = nodes.length === 1 && nodes[0].type === 'styledSpan' ? nodes[0].style?.fontSizePt : undefined;
  return fontSizePt ? fontSizePt / 11 : 1; // 11pt is Word's own default body size
}

// A rough average character advance width at Word's own 11pt default body
// size, used purely as a "no font metrics available" proxy the same way
// TXT's ASCII-table widths already approximate column sizing. Deliberately
// expressed directly in twips (not an abstract weight unit) — see
// computeContentBasedWidths's comment for why that matters.
const TWIPS_PER_CHAR_AT_11PT = 140;

function cellRequiredTwips(cell) {
  return flattenCellText(cell).length * cellFontScale(cell) * TWIPS_PER_CHAR_AT_11PT;
}

function sumRange(arr, start, count) {
  let sum = 0;
  for (let i = 0; i < count; i++) sum += arr[start + i] ?? 0;
  return sum;
}

// No explicit per-column width declared anywhere: approximate a browser's
// content-based auto table layout (which sizes each column to fit its own
// content) by computing each column's *absolute minimum* required width
// directly in twips from its longest flattened, font-size-scaled cell
// content — then distributing whatever's left over beyond those minimums
// proportionally. A flat equal split (the original behavior) badly
// squeezes a single-cell column like "Subtotal"/"Amount" next to wide
// multi-column cells in the same table, forcing mid-word wraps that never
// happen in the browser.
//
// Minimums are computed in real twips *up front*, not as a relative "weight
// unit" later normalized against every other column's weight — that
// distinction matters once colspan'd cells are involved: a first attempt at
// this heuristic worked in abstract weight units and only asked "does this
// span already have enough weight, relative to what's already there?" —
// but an unrelated large cell elsewhere in the same table (e.g. a long
// company address) inflates the *total* weight pool, diluting everyone
// else's proportional share even when their own weight nominally "met" its
// requirement. Working in absolute twips from the start and only doing
// proportional math on the *leftover* space avoids that dilution — which is
// exactly what let a `font-size:60px` "INVOICE" banner cell still end up too
// narrow to fit on one line despite this function's own deficit-boosting.
//
// Colspan'd cells are handled in two passes: first every colspan===1 cell
// sets its column's baseline minimum directly; then each colspan'd cell
// checks whether its *spanned columns' combined* minimum is already enough
// to fit its own content — if not, the shortfall is added evenly across
// those columns.
function computeContentBasedWidths(grid, columnCount, availableWidthTwips) {
  const minTwips = new Array(columnCount).fill(0);
  const spanningCells = [];

  for (const gridRow of grid) {
    let col = 0;
    for (const slot of gridRow) {
      if (!slot.continuation && slot.cell) {
        if (slot.colspan === 1) {
          minTwips[col] = Math.max(minTwips[col], cellRequiredTwips(slot.cell));
        } else {
          spanningCells.push({ col, colspan: slot.colspan, cell: slot.cell });
        }
      }
      col += slot.colspan;
    }
  }

  for (const { col, colspan, cell } of spanningCells) {
    const deficit = cellRequiredTwips(cell) - sumRange(minTwips, col, colspan);
    if (deficit > 0) {
      const share = deficit / colspan;
      for (let i = 0; i < colspan; i++) minTwips[col + i] += share;
    }
  }

  // A pure spacer column (no content anywhere) still gets a small minimum
  // width rather than collapsing toward zero (~0.25in).
  const MIN_COL_TWIPS = 360;
  const floored = minTwips.map((w) => Math.max(MIN_COL_TWIPS, w));
  const totalMin = floored.reduce((a, b) => a + b, 0);

  if (totalMin >= availableWidthTwips) {
    // The table's own minimum content doesn't fit at all — scale every
    // column down proportionally so it still sums to exactly the page's
    // content width, since fixed table layout requires that.
    const scale = availableWidthTwips / totalMin;
    return floored.map((w) => Math.max(1, Math.floor(w * scale)));
  }

  // Distribute whatever's left over beyond everyone's minimum proportionally
  // to each column's own minimum, so a column that already needed more room
  // also gets a proportionally bigger share of any leftover space, rather
  // than every column growing by the same flat amount. `floored`/deficit
  // shares can carry fractional twips at this point (e.g. a deficit split
  // across an odd colspan) — rounded here since w:w must be a plain integer,
  // not the invalid decimal string a stray fractional value would produce.
  const slack = availableWidthTwips - totalMin;
  const widths = floored.map((w) => Math.round(w + (w / totalMin) * slack));
  const used = widths.reduce((a, b) => a + b, 0);
  widths[widths.length - 1] += availableWidthTwips - used; // absorb any rounding leftover
  return widths;
}

// expandTableGrid (grid[row] of slots, each a real cell or a vMerge/hMerge
// continuation carrying `originCell`) now lives in shared/table-grid.js —
// PPTX's own table support needs the identical row×column expansion.

const TWIPS_PER_PX = 15; // 1440 twips/in ÷ 96 CSS px/in

function renderTableRow(gridRow, isHeaderRow, colWidths, rctx) {
  const b = new XmlBuilder();
  b.el('w:tr');
  if (gridRow.heightPx) {
    // A row that carries no real content (a decorative colored strip, just
    // &nbsp;) needs hRule="exact" to actually shrink to the declared height —
    // "atLeast" still leaves room for an empty paragraph's default line
    // height, which is what made a 4px CSS strip render as a thick bar. A
    // row with real content keeps "atLeast" so text is never clipped.
    const hRule = gridRow.every((slot) => slot.continuation || cellIsEmpty(slot.cell)) ? 'exact' : 'atLeast';
    b.el('w:trPr');
    b.self('w:trHeight', { 'w:val': Math.round(gridRow.heightPx * TWIPS_PER_PX), 'w:hRule': hRule });
    b.up();
  }
  let col = 0;
  for (const slot of gridRow) {
    const width = colWidths.slice(col, col + slot.colspan).reduce((sum, w) => sum + w, 0);
    b.raw(renderTableCellSlot(slot, isHeaderRow, width, rctx));
    col += slot.colspan;
  }
  b.up();
  return b.end();
}

function cellIsEmpty(cell) {
  return !cell || (cell.children ?? []).length === 0;
}

// A per-cell border (e.g. border-bottom:1px solid #hex used as a divider
// line under a decorative row) is independent of the table's own `border`
// (see the table.border branch in renderTable) — a cell can have one even
// when the table has none.
function renderCellBorders(b, cell) {
  if (!cell) return;
  // An explicit `border:none` must actively suppress whatever border the
  // table's own style (e.g. TableGrid on a real data table) would otherwise
  // draw for this cell — w:val="nil" is OOXML's "no border, override the
  // table style's default" (as opposed to omitting w:tcBorders entirely,
  // which means "no override, inherit the table style's border here").
  if (cell.noBorder) {
    b.el('w:tcBorders');
    for (const edge of BOX_EDGES) b.self(`w:${edge}`, { 'w:val': 'nil' });
    b.up();
    return;
  }
  const edgeBorders = [['top', cell.borderTop], ['left', cell.borderLeft], ['bottom', cell.borderBottom], ['right', cell.borderRight]];
  if (!edgeBorders.some(([, border]) => border)) return;
  b.el('w:tcBorders');
  for (const [edge, border] of edgeBorders) {
    if (border) b.self(`w:${edge}`, { 'w:val': border.style, 'w:sz': border.sizeEighthPt, 'w:color': border.color });
  }
  b.up();
}

// A cell's own `padding` (e.g. `th,td{padding:8px}`) — OOXML's per-cell
// equivalent is w:tcMar, a direct "space between the cell border and its
// content" on all four sides, unlike a plain paragraph/heading's padding
// (see resolveEdgeBorders' `invisible` synthesized border) which has no such
// dedicated construct to begin with. Omitted entirely when the cell
// declared no padding, leaving Word's own default cell margins untouched.
function renderTcMar(b, padding) {
  if (!padding) return;
  b.el('w:tcMar');
  for (const edge of BOX_EDGES) {
    const pt = padding[edge];
    if (pt === undefined) continue;
    b.self(`w:${edge}`, { 'w:w': Math.round(pt * 20), 'w:type': 'dxa' });
  }
  b.up();
}

function renderTableCellSlot(slot, isHeaderRow, width, rctx) {
  const b = new XmlBuilder();
  b.el('w:tc');
  b.el('w:tcPr');
  b.self('w:tcW', { 'w:w': width, 'w:type': 'dxa' });
  if (slot.colspan > 1) b.self('w:gridSpan', { 'w:val': slot.colspan });
  if (slot.continuation) b.self('w:vMerge');
  else if (slot.restart) b.self('w:vMerge', { 'w:val': 'restart' });
  // A continuation slot has no `cell` of its own (see expandTableGrid) — its
  // border/background/padding come from `originCell`, the rowspan's real
  // cell, so the merged region looks continuous instead of the covered rows
  // falling back to the table's default grid.
  const styleSource = slot.cell ?? slot.originCell;
  // CT_TcPr is a strict, order-sensitive schema — tcBorders, shd, tcMar,
  // textDirection must appear in exactly this sequence or Word may silently
  // disregard the out-of-order element (or flag the file for repair).
  renderCellBorders(b, styleSource);
  if (styleSource?.backgroundColor) b.self('w:shd', { 'w:val': 'clear', 'w:fill': styleSource.backgroundColor });
  renderTcMar(b, styleSource?.padding);
  // `writing-mode:vertical-rl` — Word's closest equivalent is a top-to-bottom,
  // right-to-left column flow, the same reading direction CSS's vertical-rl
  // describes; arbitrary rotation angles (CSS `transform:rotate()`) have no
  // table-cell equivalent in WordprocessingML and are left unrepresented.
  if (styleSource?.writingMode === 'vertical-rl') b.self('w:textDirection', { 'w:val': 'tbRl' });
  b.up();

  b.raw(slot.continuation ? '<w:p/>' : renderCellContent(slot.cell, isHeaderRow, width, rctx));

  b.up();
  return b.end();
}

// A cell's content is either plain inline runs (the common case — wrap in one
// paragraph) or block content (multiple paragraphs, a nested table — see the
// "Additions since the original plan" note: TableCellNode.children can hold
// either since the shared parser fix). Unlike ASCII/GFM output, a real <w:tc>
// natively supports multiple <w:p> paragraphs and a nested <w:tbl>, so no
// flattening is needed either way. `cell.align` (explicit text-align, or the
// browser-default center for a <th> — see dom-to-idm.js's mapCells) applies
// to the plain-inline case directly, and is the fallback for a block child
// (e.g. a <p> inside the cell) that doesn't declare its own alignment.
function renderCellContent(cell, isHeaderRow, width, rctx) {
  const nodes = cell.children ?? [];
  const bold = isHeaderRow || cell.header;
  const rtl = cell.direction === 'rtl';
  if (nodes.length === 0) return '<w:p/>'; // OOXML requires at least one paragraph per cell
  if (nodes.every((n) => !isBlockNode(n))) return renderCellParagraph(nodes, bold, rctx, cell.align, rtl, width);
  return nodes.map((n) => renderCellBlock(n, bold, width, rctx, cell.align, rtl)).join('') || '<w:p/>';
}

// A cell's own actual width (twips, incl. colspan — see renderTableRow) — a
// bare `<img>` cell (this project's own "Preview" column) reaches an
// `<w:r>` run through this exact paragraph, and a run has no width of its
// own to clamp against; `maxWidthTwips` rides along in `rProps` (the same
// mechanism bold/italic/etc. already thread through nested inline nodes)
// purely so computeImageExtent (via renderImageRun) can clamp against the
// CELL's real width instead of always the full page content width. Minus a
// default cell margin per side (LAYOUT_CELL_MARGIN_TWIPS) — a cell's
// content area is narrower than its own declared grid width by that much,
// same inset renderLayoutRow's own shape-width computation already
// subtracts for the identical reason.
function renderCellParagraph(inlineNodes, bold, rctx, align, rtl, widthTwips) {
  const b = new XmlBuilder();
  b.el('w:p');
  if (align || rtl) {
    b.el('w:pPr');
    if (align) b.self('w:jc', { 'w:val': align });
    // `direction:rtl` — marks the paragraph mark itself as bidirectional, the
    // paragraph-level counterpart to each run's own w:rtl (see renderRPr).
    if (rtl) b.self('w:bidi');
    b.up();
  }
  const maxWidthTwips = widthTwips != null ? Math.max(1, widthTwips - 2 * LAYOUT_CELL_MARGIN_TWIPS) : undefined;
  b.raw(renderRuns(inlineNodes, { ...(bold ? { bold: true } : {}), ...(rtl ? { rtl: true } : {}), maxWidthTwips }, rctx));
  return b.end();
}

function renderCellBlock(node, bold, width, rctx, cellAlign, rtl) {
  if (node.type === 'paragraph') return renderCellParagraph(node.children, bold, rctx, node.align ?? cellAlign, rtl, width);
  if (node.type === 'table') return renderTable(node, rctx, width);
  return renderBlock(node, rctx);
}


// True only if the table has at least one genuinely non-empty cell AND no
// cell holds block content — a decorative all-empty spacer table and a
// paragraph/nested-table-holding layout table both return false here.
function isDataTable(rows) {
  let hasContent = false;
  for (const row of rows) {
    for (const cell of row.cells) {
      const nodes = cell.children ?? [];
      if (nodes.some(isBlockNode)) return false;
      if (nodes.length > 0) hasContent = true;
    }
  }
  return hasContent;
}

// A per-cell style field this project's own mapCells (dom-to-idm.js) sets
// from CSS — presence of ANY of these on ANY cell means the source HTML's
// author deliberately styled this table, just not with a visible border.
// `align` is deliberately excluded here — see cellHasExplicitStyle below,
// which handles it separately since it isn't purely CSS-driven.
const CELL_STYLE_FIELDS = [
  'backgroundColor', 'borderTop', 'borderRight', 'borderBottom', 'borderLeft', 'noBorder',
  'padding', 'verticalAlign', 'color', 'bold', 'italic', 'underline', 'strikethrough',
  'fontSizePt', 'direction', 'writingMode', 'rotateDeg', 'widthPercent',
];

// `cell.align` alone isn't a reliable "the author styled this" signal — a
// <th> gets `align: 'center'` by browser-default convention even with zero
// CSS at all (mapCells: `extractAlign(...) ?? (isHeader ? 'center' : undefined)`),
// which would otherwise make an entirely bare `<table><th>...` table wrongly
// look "explicitly styled" and lose its sensible bordered-TableGrid default.
// A real `text-align` declaration on a header cell is indistinguishable from
// that convention when it also happens to be 'center' — an acceptable false
// negative, since visually there's no difference either way.
function cellHasExplicitStyle(cell) {
  if (CELL_STYLE_FIELDS.some((field) => cell[field] !== undefined)) return true;
  return cell.align !== undefined && !(cell.header && cell.align === 'center');
}

// isDataTable's default (falling through to Word's bordered "TableGrid"
// style below) assumes an author who declared literally zero CSS anywhere —
// the common case for content copy-pasted with no styling at all, where a
// visible grid is the sensible Word-side default. A table the author DID
// style in some way (e.g. this project's own `<table style="width:50%;
// margin-left:auto">` with per-cell `text-align:right`, but deliberately no
// border) should have that specific choice — no border — respected rather
// than treated as an oversight, the same "explicit CSS beats inference"
// rule this file already applies to table.border itself.
function tableHasExplicitStyling(table, rows) {
  if (table.widthPercent != null || table.marginAuto != null || table.align != null) return true;
  for (const row of rows) {
    for (const cell of row.cells) {
      if (cellHasExplicitStyle(cell)) return true;
    }
  }
  return false;
}

// ─── Images (W8) ────────────────────────────────────────────────────────────
// resolveAssets() (called in index.js, before this ever runs) fetches real
// bytes and stamps ImageNode.assetId; rctx.images maps that assetId to the
// relationship id already allocated for its word/media/ file. This renders a
// real <w:drawing>/<a:blip r:embed>, unlike the prior attempt, which always
// fell back to a "[Image: alt]" text placeholder because its relationship
// list was never actually populated.

const EMU_PER_PX = 9525; // 914400 EMU/in ÷ 96 CSS px/in
const MAX_IMAGE_CX = CONTENT_WIDTH_TWIPS * 635; // twips → EMU (1 twip = 635 EMU); matches page content width

// resolveImagePxSize resolves the image's own real intrinsic pixel size
// (falling back to a declared HTML width/height only for an undecodable
// format — see that function's own doc comment). Either way the result is
// scaled down (preserving aspect ratio) if it would overflow `maxCx` — the
// actual CELL's content width when this image sits inside a table cell (see
// renderCellParagraph's own maxWidthTwips), or the full page content width
// for a body-level image (the default, unchanged from before). Without a
// cell-aware cap, a real (often much larger than any CSS-declared size)
// photo inside a narrow "Preview"-style column overflowed straight past its
// own cell — Word doesn't shrink an inline picture to fit a fixed-width
// cell, it just lets the picture visually bleed into whatever sits next to
// it, reading as the image being "cut off" at the cell boundary.
function computeImageExtent(imageNode, asset, maxWidthTwips) {
  const { widthPx, heightPx } = resolveImagePxSize(imageNode, asset);
  let cx = widthPx * EMU_PER_PX;
  let cy = heightPx * EMU_PER_PX;
  const maxCx = maxWidthTwips != null ? Math.min(MAX_IMAGE_CX, maxWidthTwips * 635) : MAX_IMAGE_CX;
  if (cx > maxCx) {
    const scale = maxCx / cx;
    cx = Math.round(cx * scale);
    cy = Math.round(cy * scale);
  }
  return { cx, cy };
}

function renderImageParagraph(node, rctx) {
  const b = new XmlBuilder();
  b.el('w:p');
  b.el('w:pPr');
  b.self('w:jc', { 'w:val': 'center' });
  b.up();
  b.raw(renderImageRun(node, rctx));
  return b.end();
}

function renderFigure(node, rctx) {
  let out = renderImageParagraph(node.image, rctx);
  if (node.caption) {
    const b = new XmlBuilder();
    b.el('w:p');
    b.el('w:pPr');
    b.self('w:jc', { 'w:val': 'center' });
    b.up();
    b.raw(textRun(node.caption, { italic: true }));
    out += b.end();
  }
  return out;
}

// Unresolved (fetch failed, no assetId) images render nothing — resolveAssets
// already emitted a ctx.warn for the failure, so this doesn't warn again.
function renderImageRun(imageNode, rctx, maxWidthTwips) {
  const asset = rctx.images?.get(imageNode.assetId);
  if (!asset) return '';

  const docPrId = rctx.nextImageId();
  const name = `Image${docPrId}`;
  const { cx, cy } = computeImageExtent(imageNode, asset, maxWidthTwips);

  const b = new XmlBuilder();
  b.el('w:r');
  b.el('w:drawing');
  b.el('wp:inline', { distT: 0, distB: 0, distL: 0, distR: 0 });
  b.self('wp:extent', { cx, cy });
  b.self('wp:effectExtent', { l: 0, t: 0, r: 0, b: 0 });
  b.self('wp:docPr', { id: docPrId, name });
  b.el('wp:cNvGraphicFramePr');
  b.self('a:graphicFrameLocks', { 'xmlns:a': NS.a, noChangeAspect: 1 });
  b.up(); // wp:cNvGraphicFramePr

  b.el('a:graphic', { 'xmlns:a': NS.a });
  b.el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture' });
  b.el('pic:pic', { 'xmlns:pic': NS.pic });

  b.el('pic:nvPicPr');
  b.self('pic:cNvPr', { id: docPrId, name });
  b.self('pic:cNvPicPr');
  b.up(); // pic:nvPicPr

  b.el('pic:blipFill');
  b.self('a:blip', { 'r:embed': asset.relId });
  b.el('a:stretch');
  b.self('a:fillRect');
  b.up(); // a:stretch
  b.up(); // pic:blipFill

  b.el('pic:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x: 0, y: 0 });
  b.self('a:ext', { cx, cy });
  b.up(); // a:xfrm
  b.el('a:prstGeom', { prst: 'rect' });
  b.self('a:avLst');
  b.up(); // a:prstGeom
  b.up(); // pic:spPr

  return b.end(); // auto-closes pic:pic, a:graphicData, a:graphic, wp:inline, w:drawing, w:r
}

// ─── SVG vector shapes ──────────────────────────────────────────────────────
// One SVG (`svgShapes`, see dom-to-idm.js's extractSvgShapes) needs several
// separately-positioned DrawingML shapes composited at their correct
// relative offsets (e.g. a circle centered inside a rect) — unlike a single
// picture/card, `wp:inline` can't do this: sequential inline drawings just
// stack top-to-bottom in reading order, with no way to overlay one on
// another. Each shape is instead drawn as its own floating `wp:anchor`
// (a real ECMA-376 element — not the `wpg:wgp` group-shape extension, which
// this project has no way to verify against a real Word instance), all
// sharing ONE host paragraph and both positioned `relativeFrom="paragraph"`/
// `"column"` — since every shape's offset is measured from that same shared
// origin, they end up composited exactly like the source SVG, the same way
// Word's own floating-image positioning already works for any document.
function svgShapeAnchorXml(shape, rctx, scale) {
  const geo = resolveSvgShapeGeometry(shape, scale);
  const cx = Math.max(1, Math.round(geo.cx));
  const cy = Math.max(1, Math.round(geo.cy));
  // relativeHeight (the anchor's required z-order value) just needs to be
  // unique per anchor in the document — reusing the same monotonic counter
  // docPr ids already come from is a convenient, always-unique source.
  const docPrId = rctx.nextImageId();

  const b = new XmlBuilder();
  b.el('w:r');
  b.el('w:drawing');
  b.el('wp:anchor', {
    behindDoc: 0, distT: 0, distB: 0, distL: 0, distR: 0, simplePos: 0,
    locked: 0, layoutInCell: 1, allowOverlap: 1, relativeHeight: docPrId,
  });
  b.self('wp:simplePos', { x: 0, y: 0 });
  b.el('wp:positionH', { relativeFrom: 'column' });
  b.el('wp:posOffset');
  b.text(String(Math.round(geo.x)));
  b.up();
  b.up();
  b.el('wp:positionV', { relativeFrom: 'paragraph' });
  b.el('wp:posOffset');
  b.text(String(Math.round(geo.y)));
  b.up();
  b.up();
  b.self('wp:extent', { cx, cy });
  b.self('wp:effectExtent', { l: 0, t: 0, r: 0, b: 0 });
  b.self('wp:wrapNone');
  b.self('wp:docPr', { id: docPrId, name: `Shape ${docPrId}` });
  b.self('wp:cNvGraphicFramePr');
  b.el('a:graphic', { 'xmlns:a': NS.a });
  b.el('a:graphicData', { uri: NS_WPS });
  b.el('wps:wsp', { 'xmlns:wps': NS_WPS });
  b.self('wps:cNvSpPr');
  b.el('wps:spPr');
  b.el('a:xfrm', { flipH: geo.flipH ? 1 : undefined });
  b.self('a:off', { x: 0, y: 0 });
  b.self('a:ext', { cx, cy });
  b.up();
  if (geo.prst === 'custom') {
    b.el('a:custGeom');
    b.el('a:pathLst');
    b.el('a:path', { w: cx, h: cy });
    geo.points.forEach(([px, py], i) => {
      b.el(i === 0 ? 'a:moveTo' : 'a:lnTo');
      b.self('a:pt', { x: px, y: py });
      b.up();
    });
    if (geo.closed) b.self('a:close');
    b.up(); // a:path
    b.up(); // a:pathLst
    b.up(); // a:custGeom
  } else {
    b.el('a:prstGeom', { prst: geo.prst });
    b.self('a:avLst');
    b.up();
  }
  b.raw(shapeFillXml({ backgroundColor: shape.fill }));
  if (shape.stroke) b.raw(shapeLineXml({ color: shape.stroke, sizeEighthPt: Math.max(2, Math.round((shape.strokeWidthPx ?? 1) * 8)) }));
  return b.end(); // auto-closes wps:wsp, a:graphicData, a:graphic, wp:anchor, w:drawing, w:r
}

function renderSvgShapesParagraph(node, rctx) {
  const naturalCx = node.widthPx * EMU_PER_PX;
  const naturalCy = node.heightPx * EMU_PER_PX;
  let cx = naturalCx;
  let cy = naturalCy;
  if (cx > MAX_IMAGE_CX) {
    const scale = MAX_IMAGE_CX / cx;
    cx = Math.round(cx * scale);
    cy = Math.round(cy * scale);
  }
  const scale = cx / node.widthPx; // EMU per SVG-unit — uniform, aspect always preserved

  const b = new XmlBuilder();
  b.el('w:p');
  // A floating (anchored) drawing doesn't reserve its own layout height the
  // way an inline one does — Word just overlays it on top of the page at
  // its own paragraph's line position. Without an explicit w:after spacing
  // reserving the SVG's own height, whatever paragraph comes next would
  // start immediately below this (visually empty) host paragraph's single
  // line, overlapping the floating shapes instead of flowing below them.
  b.el('w:pPr');
  renderSpacing(b, { afterPt: cy / EMU_PER_PT });
  b.up();
  b.raw(node.shapes.map((shape) => svgShapeAnchorXml(shape, rctx, scale)).join(''));
  return b.end();
}

// ─── Inline (run) rendering ─────────────────────────────────────────────────
// rProps accumulates formatting through nested inline nodes (e.g.
// <strong><em>x</em></strong>) so it can all be applied to the leaf run.

function renderRuns(nodes, rProps, rctx) {
  return nodes.map((n) => renderRun(n, rProps, rctx)).join('');
}

function renderRun(node, rProps, rctx) {
  switch (node.type) {
    case 'text':
      return textRun(node.value, rProps);
    case 'strong':
      return renderRuns(node.children, { ...rProps, bold: true }, rctx);
    case 'emphasis':
      return renderRuns(node.children, { ...rProps, italic: true }, rctx);
    case 'codeSpan':
      return textRun(node.value, { ...rProps, code: true });
    case 'styledSpan':
      return renderRuns(node.children, { ...rProps, ...node.style }, rctx);
    case 'link':
      return renderHyperlink(node, rProps, rctx);
    case 'lineBreak':
      return '<w:r><w:br/></w:r>';
    case 'image':
      return renderImageRun(node, rctx, rProps.maxWidthTwips);
    default:
      return '';
  }
}

// Hyperlink relationships are External (they point at a URL, not another
// package part) — the id is allocated through the same RelationshipManager
// used for styles/numbering/images, so document.xml and .rels can't drift.
// An href starting with "#" is an in-document fragment link (e.g.
// <a href="#end"> pointing at <... id="end">), not a real external URL —
// registering it as an External relationship would produce a dead link with
// a nonsensical target. It links via w:anchor to a bookmark instead (see
// renderBookmarked, which wraps whatever element carried the matching id).
function renderHyperlink(node, rProps, rctx) {
  const b = new XmlBuilder();
  if (node.href.startsWith('#')) {
    b.el('w:hyperlink', { 'w:anchor': bookmarkName(node.href.slice(1)) });
  } else {
    const rId = rctx.relationships.add(REL_TYPES.hyperlink, node.href, 'External');
    b.el('w:hyperlink', { 'r:id': rId });
  }
  b.raw(renderRuns(node.children, { ...rProps, hyperlink: true }, rctx));
  return b.end();
}

// OOXML bookmark names must start with a letter/underscore and contain only
// word characters (no spaces, punctuation, etc.) — an HTML id can contain
// characters a bookmark name can't, so anything else is replaced.
function bookmarkName(id) {
  const cleaned = id.replace(/[^A-Za-z0-9_]/g, '_');
  return /^[A-Za-z_]/.test(cleaned) ? cleaned : `_${cleaned}`;
}

// Wraps a rendered paragraph/heading with a bookmark when its source element
// carried an id="..." — the counterpart to renderHyperlink's w:anchor.
function renderBookmarked(node, xml, rctx) {
  if (!node.anchorId) return xml;
  const id = rctx.nextBookmarkId();
  const name = bookmarkName(node.anchorId);
  return `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${xml}<w:bookmarkEnd w:id="${id}"/>`;
}

function textRun(value, rProps = {}) {
  const b = new XmlBuilder();
  b.el('w:r');
  b.raw(renderRPr(rProps));
  b.el('w:t', { 'xml:space': 'preserve' });
  b.text(value);
  const run = b.end();
  return rProps.formField ? wrapFormField(run) : run;
}

// An <input>/<textarea>/<select>'s value — real, native Word form content:
// a "Plain Text" content control (`w:sdt`), the same construct Word's own
// Developer tab inserts, rendered with a real shaded/bordered fillable-field
// box. Unlike a table (the trick used for `<button>` — see
// renderShrinkToFitBox), a content control sits inline inside a run, so it
// can appear mid-paragraph next to a label, not just as its own block.
// `w:id` only needs to not collide within this one document — an
// incrementing counter is enough, real uniqueness across documents isn't
// required for Word to render it correctly.
let nextFormFieldId = 100000;

function wrapFormField(runXml) {
  const b = new XmlBuilder();
  b.el('w:sdt');
  b.el('w:sdtPr');
  b.self('w:id', { 'w:val': nextFormFieldId++ });
  b.self('w:text');
  b.up();
  b.el('w:sdtContent');
  b.raw(runXml);
  b.up();
  b.up();
  return b.end();
}

function renderRPr(rProps) {
  const hasAny = rProps.bold || rProps.italic || rProps.code || rProps.hyperlink ||
    rProps.underline || rProps.strikethrough || rProps.color || rProps.backgroundColor ||
    rProps.fontSizePt || rProps.verticalAlign || rProps.fontFamily || rProps.rtl;
  if (!hasAny) return '';
  const b = new XmlBuilder();
  b.el('w:rPr');
  if (rProps.hyperlink) b.self('w:rStyle', { 'w:val': 'Hyperlink' });
  if (rProps.bold) b.self('w:b');
  if (rProps.italic) b.self('w:i');
  if (rProps.underline) b.self('w:u', { 'w:val': 'single' });
  if (rProps.strikethrough) b.self('w:strike');
  // An explicit font-family wins over <code>'s default monospace font — only
  // one w:rFonts is valid per run, so this picks whichever applies. A bare
  // CSS generic keyword ("sans-serif") is mapped to a real installable name
  // first — Word has no font literally named "sans-serif" to find.
  const fontName = (rProps.fontFamily ? resolveFontFamilyName(rProps.fontFamily) : undefined) ?? (rProps.code ? 'Courier New' : undefined);
  if (fontName) b.self('w:rFonts', { 'w:ascii': fontName, 'w:hAnsi': fontName });
  if (rProps.code && !rProps.backgroundColor) b.self('w:shd', { 'w:val': 'clear', 'w:fill': 'F1F1F1' });
  if (rProps.color) b.self('w:color', { 'w:val': rProps.color });
  if (rProps.backgroundColor) b.self('w:shd', { 'w:val': 'clear', 'w:fill': rProps.backgroundColor });
  if (rProps.fontSizePt) b.self('w:sz', { 'w:val': Math.round(rProps.fontSizePt * 2) });
  if (rProps.verticalAlign) b.self('w:vertAlign', { 'w:val': rProps.verticalAlign });
  // `direction:rtl` marks the run itself as right-to-left, complementing the
  // paragraph-level w:bidi below — Word needs both to correctly lay out and
  // shape bidirectional (e.g. Arabic/Hebrew) text.
  if (rProps.rtl) b.self('w:rtl');
  return b.end();
}

// Portrait page dimensions in twips (1/1440 inch).
const PAGE_SIZES = {
  Letter: { w: 12240, h: 15840 },
  A4: { w: 11906, h: 16838 },
  Legal: { w: 12240, h: 20160 },
};
const DEFAULT_MARGIN_TWIPS = 1440; // 1in, matches the PRD's 20mm-ish default intent

function mmToTwips(mm) {
  return Math.round(mm * 56.6929); // 1440 twips/in ÷ 25.4mm/in
}

// Shared by renderSectPr (W10) and computeContentWidthTwips — both need the
// same pageSize/orientation/margins resolution, and having two independent
// copies of this logic risks them silently drifting apart (e.g. a table
// sized against one page width while w:sectPr declares a different one).
function resolvePageGeometry(options = {}) {
  const size = PAGE_SIZES[options.pageSize] ?? PAGE_SIZES.Letter;
  const landscape = options.orientation === 'landscape';
  const w = landscape ? size.h : size.w;
  const h = landscape ? size.w : size.h;

  const margins = options.margins ?? {};
  const top = margins.top !== undefined ? mmToTwips(margins.top) : DEFAULT_MARGIN_TWIPS;
  const right = margins.right !== undefined ? mmToTwips(margins.right) : DEFAULT_MARGIN_TWIPS;
  const bottom = margins.bottom !== undefined ? mmToTwips(margins.bottom) : DEFAULT_MARGIN_TWIPS;
  const left = margins.left !== undefined ? mmToTwips(margins.left) : DEFAULT_MARGIN_TWIPS;

  return { w, h, landscape, top, right, bottom, left };
}

// The real available width a top-level table has to lay out its columns in
// — without this, a table's column widths were computed against a hardcoded
// assumption (CONTENT_WIDTH_TWIPS) no matter what pageSize/orientation/
// margins the caller actually chose, so e.g. switching to landscape to fit a
// wide table did nothing for its column widths even though the page itself
// did get wider.
function computeContentWidthTwips(options) {
  const { w, left, right } = resolvePageGeometry(options);
  return w - left - right;
}

// W10 — reads convert()'s pageSize/orientation/margins options (PRD §13);
// W1's hardcoded US Letter portrait remains the default when none are given.
function renderSectPr(b, ctx) {
  const { w, h, landscape, top, right, bottom, left } = resolvePageGeometry(ctx?.options);

  b.el('w:sectPr');
  b.self('w:pgSz', landscape ? { 'w:w': w, 'w:h': h, 'w:orient': 'landscape' } : { 'w:w': w, 'w:h': h });
  b.self('w:pgMar', {
    'w:top': top, 'w:right': right, 'w:bottom': bottom, 'w:left': left,
    'w:header': 720, 'w:footer': 720, 'w:gutter': 0,
  });
  b.up();
}
