/**
 * Flattens a table cell's content into per-run text+style segments, for
 * Excel's rich-text shared-string format (<si><r><rPr>...</rPr><t>...</t>
 * </r>...</si>) — one Excel cell, several *differently* formatted runs of
 * text, matching how `<td><b>Bold</b> <i>Italic</i></td>` actually looks in
 * a browser. Plain flattenCellText/flattenNode (idm/flatten-text.js) throws
 * this per-run distinction away by design (a single plain string); this is
 * the richer sibling used specifically for XLSX string cells.
 *
 * Style is accumulated through nested <strong>/<em>/styledSpan the same way
 * DOCX's renderRuns/renderRun accumulates rProps through nesting — a run's
 * final style is everything its ancestors contributed, closest-wins for any
 * property multiple ancestors set. Adjacent runs that end up with identical
 * style are merged, so a cell with no (or uniform) styling collapses back
 * down to a single run — the common case stays exactly as simple as before.
 */

import { flattenNode, isBlockNode } from '../../idm/flatten-text.js';
import { fontStyleKey } from './font-style-key.js';

// Points, decreasing per level — same descending-scale idea as DOCX's own
// HEADING_SIZES (there in half-points, Word's unit; here in whole points,
// this project's own fontSizePt convention).
const HEADING_FONT_SIZES = [18, 16, 14, 12, 11, 11];

function pushRun(runs, text, style) {
  if (text) runs.push({ text, ...style });
}

// Handles the inline-formatting tags that actually vary per run
// (bold/italic/color/size via strong/emphasis/styledSpan); anything else
// (a nested list/table/heading/etc., or an image) falls back to the
// existing exhaustive flattenNode logic and becomes one plain run at the
// current accumulated style — those cases don't carry meaningful *internal*
// run variation worth preserving here.
//
// A block with its own backgroundColor (e.g. `<h3 style="background:#1F4E79;
// color:white">From</h3>` sitting inside a `<td>` alongside several sibling
// <p>s — mapElement stamps the background directly on the node and wraps
// its children in one styledSpan carrying the paired `color:white`, the same
// "whole-element inline style" shape this project's own PPTX exporter keys
// off of — see pptx/shapes.js's resolveOwnFontSizePt) needs that pairing's
// *color* dropped whenever its own background can't actually become the
// containing cell's real fill — `cellBg` (the enclosing `<c>`'s own
// `backgroundColor`, threaded through every walkInline call below) is what
// tells a node whether that's the case: a plain multi-block cell has ONE
// shared fill for everything in it (worksheet.js's plain renderCell path, or
// a band that band-splitting — see worksheet.js's splitIntoBackgroundBands —
// decided NOT to give its own dedicated row/fill), so a node's own
// background silently never renders there, while the light/white text color
// chosen specifically to sit on it still would, rendering invisible against
// the cell's own (different, usually white) fill instead. Once a node HAS
// been promoted to its own single-fill band/row (cellBg === the node's own
// backgroundColor), the pairing is legitimate and must stay — that's exactly
// how "From" on its own blue band keeps its white text. Only the color is
// stripped in the mismatched case — bold/italic/fontSize the same styledSpan
// might also carry stay intact, since those don't depend on a background to
// stay legible.
function dropBackgroundPairedColor(children) {
  if (children.length !== 1 || children[0].type !== 'styledSpan') return children;
  const { color, ...restStyle } = children[0].style ?? {};
  return [{ ...children[0], style: restStyle }];
}

function walkInline(nodes, style, runs, cellBg) {
  for (const node of nodes) {
    switch (node.type) {
      case 'text':
        pushRun(runs, node.value, style);
        break;
      case 'strong':
        walkInline(node.children, { ...style, bold: true }, runs, cellBg);
        break;
      case 'emphasis':
        walkInline(node.children, { ...style, italic: true }, runs, cellBg);
        break;
      case 'styledSpan':
        walkInline(node.children, { ...style, ...node.style }, runs, cellBg);
        break;
      case 'codeSpan':
        pushRun(runs, node.value, style);
        break;
      case 'lineBreak':
        pushRun(runs, '\n', style);
        break;
      // Excel has no real per-run hyperlink relationship within a mixed cell
      // (a worksheet <hyperlink> targets a whole cell reference, not a
      // sub-range of one cell's text) — coloring the run the conventional
      // hyperlink blue+underline is the closest available approximation,
      // same category as approximating a gradient with its first color stop.
      case 'link':
        walkInline(node.children, { ...style, color: '0563C1', underline: true }, runs, cellBg);
        break;
      // A block child (most commonly a <p>) inside a cell that also holds
      // other block content (e.g. a <h2> + <p> + <ul> all in one <td>) still
      // has its own bold/italic/etc runs worth preserving — these three share
      // flattenNode's plain-text shape (just node.children, no row/item
      // structure to rebuild), so recursing keeps that styling instead of
      // collapsing straight to a plain string like the default case does.
      // See dropBackgroundPairedColor's own note on when a node.backgroundColor
      // here strips its paired text color first.
      case 'paragraph':
      case 'quote':
      case 'section':
        walkInline(node.backgroundColor && node.backgroundColor !== cellBg ? dropBackgroundPairedColor(node.children) : node.children, style, runs, cellBg);
        break;
      // A heading is conventionally bold and larger — the same "an
      // established convention counts as an explicit declaration" reasoning
      // already used for a <th>'s default center alignment elsewhere in this
      // exporter, and matching DOCX's own per-level HEADING_SIZES scale.
      case 'heading': {
        const children = node.backgroundColor && node.backgroundColor !== cellBg ? dropBackgroundPairedColor(node.children) : node.children;
        walkInline(children, { ...style, bold: true, fontSizePt: HEADING_FONT_SIZES[(node.level ?? 1) - 1] ?? 11 }, runs, cellBg);
        break;
      }
      // A flattened list still reads better as one marker-prefixed line per
      // item (matching how TXT/Markdown render a list that can't be a real
      // nested structure either) than run together as "Item1 Item2" with no
      // separation at all — the plain flattenNode fallback below does the
      // latter, which is fine for CSV/DOCX's column-width heuristic (no
      // visual difference matters there) but reads badly as actual cell text.
      case 'list':
        (node.items ?? []).forEach((item, i) => {
          if (i > 0) pushRun(runs, '\n', style);
          pushRun(runs, node.ordered ? `${i + (node.start ?? 1)}. ` : '• ', style);
          walkInline(item.children, style, runs, cellBg);
        });
        break;
      default:
        pushRun(runs, flattenNode(node), style);
    }
  }
}

function mergeAdjacent(runs) {
  const merged = [];
  for (const run of runs) {
    const last = merged[merged.length - 1];
    if (last && fontStyleKey(last) === fontStyleKey(run)) {
      last.text += run.text;
    } else {
      merged.push({ ...run });
    }
  }
  return merged;
}

/**
 * `cell`'s own font styling (`color`/`bold`/`italic`/`underline`/
 * `strikethrough`/`fontSizePt`, from its own style attribute — see
 * dom-to-idm.js's mapCells) is the *base* style every run
 * starts from, so `<td style="color:red"><b>A</b> B</td>` correctly renders
 * both "A" (red+bold) and " B" (red) — not just the bold run, losing the
 * cell's own color for the plain-text part. `header` is folded in the same
 * way (`<th>`/a header-row cell is always bold) because a rich shared-string
 * run with no `<rPr>` of its own renders in Excel's plain default font, not
 * the cell's own cellXfs font — unlike a plain (non-rich) string or a
 * numeric cell, where the cellXfs font is exactly what's used.
 */
export function flattenCellRuns(cell) {
  const base = {
    bold: cell.header || cell.bold,
    italic: cell.italic,
    underline: cell.underline,
    strikethrough: cell.strikethrough,
    color: cell.color,
    fontSizePt: cell.fontSizePt,
  };
  const nodes = cell.children ?? [];
  const runs = [];
  if (nodes.length > 1 && nodes.every(isBlockNode)) {
    nodes.forEach((node, i) => {
      if (i > 0) pushRun(runs, '\n', base);
      walkInline([node], base, runs, cell.backgroundColor);
    });
  } else {
    walkInline(nodes, base, runs, cell.backgroundColor);
  }
  return mergeAdjacent(runs);
}

/**
 * Same per-run flattening as flattenCellRuns, but for a top-level block node
 * (heading/paragraph/list/quote/code) rendered directly as a sheet row rather
 * than as a table cell's content — used by the XLSX exporter's whole-document
 * layout (src/exporter/xlsx/document-sheet.js). Unlike a table cell, these
 * nodes carry no cell-level `bold`/`color`/etc fields of their own (any
 * inherited color already arrives pre-wrapped as a styledSpan around
 * `node.children` — see dom-to-idm.js's mapElement) — walkInline's existing
 * `heading`/`paragraph`/`list`/`quote` cases already produce the right base
 * style (e.g. a heading's bold+size), so there's no cell-style base to seed.
 */
export function flattenBlockRuns(node) {
  const runs = [];
  // The node's own backgroundColor IS this row's own cell fill (see
  // document-sheet.js's renderBlockRow) — passing it as `cellBg` tells
  // walkInline's own heading/paragraph/etc case the pairing is legitimate
  // here (node.backgroundColor === cellBg), so it never strips the color.
  walkInline([node], {}, runs, node.backgroundColor);
  return mergeAdjacent(runs);
}

/** Trims leading/trailing whitespace across the run boundaries, dropping any run left empty. */
export function trimRuns(runs) {
  if (runs.length === 0) return runs;
  const copy = runs.map((r) => ({ ...r }));
  copy[0].text = copy[0].text.replace(/^\s+/, '');
  copy[copy.length - 1].text = copy[copy.length - 1].text.replace(/\s+$/, '');
  return copy.filter((r) => r.text !== '');
}
