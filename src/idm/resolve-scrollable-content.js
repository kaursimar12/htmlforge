/**
 * Resolves `__scrollableContainer` marker nodes (see dom-to-idm.js's
 * mapScrollableContainer) according to the `preserveScrollableContent`
 * option — the ONE place in the whole pipeline that knows this marker type
 * exists. Called once from convert() (src/index.js), between parseHtml() and
 * normalizeDocument(), so every exporter (PDF/DOCX/XLSX/PPTX) only ever sees
 * a plain, already-decided IDM tree — no exporter needs its own overflow/
 * clipping logic for this option.
 *
 * `preserveScrollableContent: true` (default) — the marker is unwrapped
 * exactly like an ordinary, non-scrollable `<div>` always has been: its
 * children replace it in place, unchanged. Output is identical to before this
 * feature existed for every caller not using the option.
 *
 * `preserveScrollableContent: false` — only the children that would fit
 * within the container's own declared height/max-height (in px) are kept,
 * in document order; the rest are dropped and a `content-clipped` warning is
 * emitted. Height is estimated per child with the exact same
 * estimateWrappedLines/font-metrics math PDF's own real pagination
 * (pdf/blocks.js) and PPTX's own real slide-fit checks already use for their
 * real layout decisions — not a separate, weaker guess invented for this
 * feature (see the plan's own note on this). Whole-child granularity only:
 * a child is either fully kept or fully dropped, never clipped mid-paragraph.
 *
 * Height also includes each block's own default vertical margin (see
 * BLOCK_MARGIN_EM below) — a real browser doesn't just spend a scrollable
 * box's height on text "ink" (line-height), it also spends it on the
 * whitespace its own default stylesheet puts around every <h1>-<h6>/<p>/
 * <ul>/<ol>/<pre>/<blockquote>. Charging only for ink used to pack roughly
 * 2x too many blocks into a declared height budget compared to what a
 * browser's own scrollable box actually shows before anyone scrolls.
 */

import { estimateWrappedLines } from '../exporter/shared/font-metrics.js';
import { flattenInline, tableRows, flattenCellText } from './flatten-text.js';
import { HEADING_SIZES_PT } from '../exporter/pdf/layout.js';

const PX_TO_PT = 0.75; // 72pt/in over 96px/in — same factor layout.js's auto-width sizing already uses.
const DEFAULT_FONT_SIZE_PT = 12;
const DEFAULT_CONTENT_WIDTH_PT = 450; // no real layout pass has run yet; a rough A4-ish content width.
// A real browser's own default body line-height ("normal") is closer to
// 1.15x font-size than the 1.2x this codebase's OTHER estimators use (e.g.
// font-metrics.js's estimateShapeSize default) — those bake in a deliberate
// safety margin against a PPTX/DOCX shape clipping its own text, which isn't
// the goal here: this estimate is trying to match what a browser's own
// scrollable viewport actually shows before scrolling, not stay safely
// oversized. This constant is local to this file, so it doesn't affect the
// separate 1.2x default the shape-sizing estimators still use.
const LINE_HEIGHT_MULTIPLIER = 1.15;

// Typical browser UA-stylesheet default vertical margin, in em of the
// block's OWN font size — e.g. Chrome/Firefox apply p{margin:1em 0}, h1{
// margin:.67em 0}, h2{margin:.83em 0}, ..., ul/ol/blockquote/pre{margin:1em
// 0}. table/li have no default vertical margin of their own in a real
// browser, so they're 0 here (a list's OWN outer margin is still charged
// once via BLOCK_MARGIN_EM.list, just not per-item).
//
// Only ONE em is charged per block, not two (top+bottom) — adjacent
// siblings' touching margins COLLAPSE in real CSS (a paragraph's bottom
// margin and the next paragraph's top margin merge into a single shared
// gap, they don't add). Charging both in full double-counted every gap
// between blocks, which under-fit a declared height budget by roughly 2x
// compared to a real browser (e.g. 200px fit only 2 paragraphs after a
// heading here, instead of the ~4-5 a real browser actually shows).
// Charging one em per block approximates that collapse: consecutive blocks
// end up sharing one gap's worth of space, the same shape as reality, even
// though a fully accurate CSS margin-collapse model (distinguishing a
// sequence's very first/last margin from the collapsed ones in between)
// isn't worth the complexity for a "roughly what's visible" estimate.
const HEADING_MARGIN_EM = [0.67, 0.83, 1, 1.33, 1.67, 2.33]; // index 0 = <h1>
const BLOCK_MARGIN_EM = { paragraph: 1, quote: 1, list: 1, code: 1, table: 0 };

function blockMarginPt(marginEm, fontSizePt) {
  return marginEm * fontSizePt;
}

export function resolveScrollableContent(document, options, warn) {
  const preserve = options?.preserveScrollableContent !== false;
  const ctx = { preserve, warn, fontSizePt: document.metadata?.fontSizePt ?? DEFAULT_FONT_SIZE_PT };
  return { ...document, children: resolveBlocks(document.children, ctx) };
}

function resolveBlocks(nodes, ctx) {
  const out = [];
  for (const node of nodes ?? []) {
    if (node.type === '__scrollableContainer') {
      out.push(...resolveContainer(node, ctx));
    } else {
      out.push(resolveNode(node, ctx));
    }
  }
  return out;
}

function resolveContainer(node, ctx) {
  const children = resolveBlocks(node.children, ctx); // resolve nested containers first
  if (ctx.preserve) return children;

  const heightPt = (node.heightPx ?? node.maxHeightPx) * PX_TO_PT;
  const kept = [];
  let usedPt = 0;
  let droppedCount = 0;
  for (const child of children) {
    const childHeightPt = estimateNodeHeightPt(child, ctx);
    if (kept.length > 0 && usedPt + childHeightPt > heightPt) {
      droppedCount++;
      continue;
    }
    kept.push(child);
    usedPt += childHeightPt;
  }
  if (droppedCount > 0) {
    ctx.warn?.('content-clipped', `Dropped ${droppedCount} block(s) from a scrollable container to fit its declared height`);
  }
  return kept;
}

// Nested containers (section/quote/list/table) recurse so a scrollable div
// buried inside a list item or table cell still gets resolved; every other
// node type (paragraph, heading, image, code, ...) is a leaf as far as this
// pass is concerned and passes through unchanged — this resolver only ever
// removes/keeps whole children, it never rewrites a node's own content.
function resolveNode(node, ctx) {
  switch (node.type) {
    case 'section':
    case 'quote':
      return { ...node, children: resolveBlocks(node.children, ctx) };
    case 'list':
      return { ...node, items: node.items.map((item) => ({ ...item, children: resolveBlocks(item.children, ctx) })) };
    case 'table': {
      const resolveRows = (rows) => rows?.map((row) => ({
        ...row,
        cells: row.cells.map((cell) => ({ ...cell, children: resolveBlocks(cell.children, ctx) })),
      }));
      return { ...node, head: resolveRows(node.head), rows: resolveRows(node.rows) };
    }
    default:
      return node;
  }
}

// Whole-child granularity, matching PPTX's estimateCardBlockHeightPt's own
// per-block-type math (shapes.js) — deliberately approximate, not a real
// layout engine (see this module's own header comment on why that's fine).
function estimateNodeHeightPt(node, ctx) {
  switch (node.type) {
    // A real <h1>-<h6> renders at its own, much larger-than-body font size
    // (HEADING_SIZES_PT — the exact sizes PDF's own renderHeading uses), not
    // the surrounding body text's size; using ctx.fontSizePt here (the old
    // behavior) undercounted both its line height AND its own, larger margin.
    case 'heading': {
      const fontSizePt = HEADING_SIZES_PT[node.level] ?? HEADING_SIZES_PT[6];
      const text = flattenInline(node.children);
      const lines = estimateWrappedLines(text, DEFAULT_CONTENT_WIDTH_PT, fontSizePt, undefined);
      const marginEm = HEADING_MARGIN_EM[(node.level ?? 1) - 1] ?? 1;
      return lines.length * fontSizePt * LINE_HEIGHT_MULTIPLIER + blockMarginPt(marginEm, fontSizePt);
    }
    case 'paragraph': {
      const text = flattenInline(node.children);
      const lines = estimateWrappedLines(text, DEFAULT_CONTENT_WIDTH_PT, ctx.fontSizePt, undefined);
      return lines.length * ctx.fontSizePt * LINE_HEIGHT_MULTIPLIER + blockMarginPt(BLOCK_MARGIN_EM.paragraph, ctx.fontSizePt);
    }
    case 'quote':
      return node.children.reduce((sum, child) => sum + estimateNodeHeightPt(child, ctx), 0) + blockMarginPt(BLOCK_MARGIN_EM.quote, ctx.fontSizePt);
    case 'section':
      return node.children.reduce((sum, child) => sum + estimateNodeHeightPt(child, ctx), 0);
    case 'list':
      return node.items.reduce((sum, item) => sum + item.children.reduce((s, child) => s + estimateNodeHeightPt(child, ctx), 0), 0) + blockMarginPt(BLOCK_MARGIN_EM.list, ctx.fontSizePt);
    case 'table':
      return tableRows(node).reduce((sum, row) => {
        const cellLines = row.cells.map((cell) => {
          const lines = estimateWrappedLines(flattenCellText(cell), DEFAULT_CONTENT_WIDTH_PT / Math.max(1, row.cells.length), ctx.fontSizePt, undefined);
          return lines.length;
        });
        return sum + Math.max(1, ...cellLines) * ctx.fontSizePt * LINE_HEIGHT_MULTIPLIER;
      }, 0);
    case 'image':
      return (node.height ?? 150) * PX_TO_PT;
    case 'code':
      return Math.max(node.value.split('\n').length, 1) * ctx.fontSizePt * LINE_HEIGHT_MULTIPLIER + blockMarginPt(BLOCK_MARGIN_EM.code, ctx.fontSizePt);
    default:
      return ctx.fontSizePt * LINE_HEIGHT_MULTIPLIER;
  }
}
