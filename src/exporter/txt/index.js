/**
 * Plain Text exporter — IDM → human-readable .txt.
 * Milestones: T1 (headings) · T2 (paragraphs) · T3 (lists) ·
 *             T4 (ASCII tables) · T5 (inline) · T6 (code/quote)
 */

import { displayImageSrc } from '../../utils/html.js';
import { isBlockNode, tableRows } from '../../idm/flatten-text.js';
import { makeListRenderer } from '../shared/list-renderer.js';

export const txtExporter = {
  id: 'txt',
  extensions: ['.txt'],
  mimeType: 'text/plain',

  export(document, ctx) {
    const baseUrl = ctx?.options?.baseUrl;
    const lines = renderBlocks(document.children, 0, baseUrl);
    return Buffer.from(lines.trimEnd() + '\n', 'utf8');
  },
};

// ─── Block rendering ──────────────────────────────────────────────────────────
// Every block renderer returns its own line(s) ending with exactly ONE trailing
// '\n'. renderBlocks joins blocks with an extra '\n', so two adjacent blocks
// end up separated by exactly one blank line.
// `baseUrl` (from convert()'s own options) is threaded through purely so it
// can reach an `image` node wherever one is nested — see renderImage.

function renderBlocks(nodes, depth, baseUrl) {
  return nodes
    .map((n) => renderBlock(n, depth, baseUrl))
    .filter((s) => s.trim() !== '')
    .join('\n');
}

function renderBlock(node, depth, baseUrl) {
  switch (node.type) {
    case 'heading':      return renderHeading(node, baseUrl);
    case 'paragraph':    return renderInline(node.children, baseUrl) + '\n';
    case 'list':         return renderList(node, depth, baseUrl);
    case 'table':         return renderTable(node, baseUrl);
    case 'code':          return renderCode(node);
    case 'quote':         return renderQuote(node, baseUrl);
    case 'image':          return renderImage(node, baseUrl) + '\n';
    case 'figure':          return renderFigure(node, baseUrl);
    case 'thematicBreak':
    case 'pageBreak':        return '-'.repeat(40) + '\n';
    case 'section':
    case 'footnote':          return renderBlocks(node.children, depth, baseUrl);
    default:                   return '';
  }
}

// T1 — headings
function renderHeading(node, baseUrl) {
  const label = renderInline(node.children, baseUrl);
  if (node.level === 1) {
    const upper = label.toUpperCase();
    return upper + '\n' + '='.repeat(upper.length) + '\n';
  }
  if (node.level === 2) {
    return label + '\n' + '-'.repeat(label.length) + '\n';
  }
  return label + '\n';
}

// T3 — lists
const { renderList } = makeListRenderer(renderInline, renderBlock);

// T4 — tables
// A <table> used for real tabular data (short inline values per cell) renders
// as an ASCII grid. A <table> used for layout/grouping (a common pattern in
// HTML emails/CMS exports) has cells holding block content — paragraphs,
// nested tables — instead; cramming that onto single grid-cell lines is
// unreadable, so it renders as normal flowing text instead. Both branches use
// the same block-vs-inline signal the parser already computes per cell.
function renderTable(table, baseUrl) {
  if (isEmptyTable(table)) return ''; // e.g. a decorative <table><tr><td>&nbsp;</td></tr></table> divider
  return hasBlockCell(table) ? renderTableAsFlow(table, baseUrl) : renderTableAsGrid(table, baseUrl);
}

function isEmptyTable(table) {
  return tableRows(table).every((row) => row.cells.every((cell) => (cell.children ?? []).length === 0));
}

function hasBlockCell(table) {
  return tableRows(table).some((row) => row.cells.some((cell) => (cell.children ?? []).some(isBlockNode)));
}

function renderTableAsGrid(table, baseUrl) {
  const allRows = tableRows(table);
  if (allRows.length === 0) return '';

  const colCount = Math.max(...allRows.map((r) => r.cells.length));
  const grid = allRows.map((row) =>
    Array.from({ length: colCount }, (_, i) => tableCellText(row.cells[i], baseUrl)),
  );
  const widths = Array.from({ length: colCount }, (_, i) =>
    Math.max(...grid.map((row) => row[i].length)),
  );

  const border = '+' + widths.map((w) => '-'.repeat(w + 2)).join('+') + '+';
  const rowLine = (cells) =>
    '|' + cells.map((c, i) => ` ${c.padEnd(widths[i])} `).join('|') + '|';

  const headCount = (table.head ?? []).length;
  const lines = [border];
  grid.forEach((cells, i) => {
    lines.push(rowLine(cells));
    if (headCount > 0 && i === headCount - 1) lines.push(border);
  });
  lines.push(border);

  return lines.join('\n') + '\n';
}

// Grid cells never hold block content (that routes to renderTableAsFlow
// instead), so this only ever needs to handle plain inline runs.
function tableCellText(cell, baseUrl) {
  if (!cell) return '';
  return renderInline(cell.children ?? [], baseUrl).replace(/\r?\n/g, ' ');
}

function renderTableAsFlow(table, baseUrl) {
  const allRows = tableRows(table);
  const sections = [];
  for (const row of allRows) {
    for (const cell of row.cells) {
      const nodes = cell.children ?? [];
      if (nodes.length === 0) continue;
      const text = nodes.every((n) => !isBlockNode(n)) ? renderInline(nodes, baseUrl) + '\n' : renderBlocks(nodes, 0, baseUrl);
      if (text.trim() !== '') sections.push(text.trimEnd());
    }
  }
  return sections.length ? sections.join('\n\n') + '\n' : '';
}

// T6 — code blocks and blockquotes
function renderCode(node) {
  return node.value.split('\n').map((line) => '    ' + line).join('\n') + '\n';
}

function renderQuote(node, baseUrl) {
  const inner = renderBlocks(node.children, 0, baseUrl).trimEnd();
  return inner.split('\n').map((line) => '> ' + line).join('\n') + '\n';
}

// Images / figures (block context)
// A data: URI src (megabytes of base64) is never printed — treated the same
// as having no src at all, falling through to alt-only or the bare marker.
function renderImage(node, baseUrl) {
  const alt = node.alt || '';
  const src = displayImageSrc(node.src, baseUrl) || '';
  if (alt && src) return `[image: ${alt} (${src})]`;
  if (alt) return `[image: ${alt}]`;
  if (src) return `[image: ${src}]`;
  return '[image]';
}

function renderFigure(node, baseUrl) {
  const img = renderImage(node.image, baseUrl);
  return node.caption ? img + '\n' + node.caption + '\n' : img + '\n';
}

// ─── Inline rendering (T5) ───────────────────────────────────────────────────
// Plain text has no way to represent bold/italic, so strong/emphasis just
// unwrap to their inner text.

function renderInline(nodes, baseUrl) {
  return nodes.map((n) => renderInlineNode(n, baseUrl)).join('');
}

// Cells that reach here (grid mode, or the plain-inline branch of flow mode)
// are guaranteed to hold only inline nodes — block content is handled by
// renderTableAsFlow before renderInline is ever called on it.
function renderInlineNode(node, baseUrl) {
  switch (node.type) {
    case 'text':      return node.value;
    case 'strong':
    case 'emphasis':
    case 'styledSpan': return renderInline(node.children, baseUrl);
    case 'codeSpan':  return `\`${node.value}\``;
    case 'link':      return `${renderInline(node.children, baseUrl)} (${node.href})`;
    case 'lineBreak': return '\n';
    case 'image':     return renderImage(node, baseUrl);
    default:          return '';
  }
}
