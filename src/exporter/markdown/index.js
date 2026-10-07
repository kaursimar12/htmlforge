/**
 * Markdown exporter — IDM → GitHub Flavored Markdown.
 * Milestones: M1 (headings/paragraphs) · M2 (inline) · M3 (lists) ·
 *             M4 (tables) · M5 (code/quote) · M6 (images/figures) · M7 (breaks)
 */

import { displayImageSrc } from '../../utils/html.js';
import { isBlockNode, tableRows } from '../../idm/flatten-text.js';
import { makeListRenderer } from '../shared/list-renderer.js';

export const markdownExporter = {
  id: 'markdown',
  extensions: ['.md', '.markdown'],
  mimeType: 'text/markdown',

  export(document, ctx) {
    const baseUrl = ctx?.options?.baseUrl;
    const lines = renderBlocks(document.children, 0, baseUrl);
    return Buffer.from(lines.trimEnd() + '\n', 'utf8');
  },
};

// ─── Block rendering ──────────────────────────────────────────────────────────
// `baseUrl` (from convert()'s own options) is threaded through every block/
// inline renderer purely so it can reach an `image` node wherever one is
// nested — it never affects any other rendering decision.

function renderBlocks(nodes, depth, baseUrl) {
  return nodes.map((n) => renderBlock(n, depth, baseUrl)).join('\n');
}

function renderBlock(node, depth, baseUrl) {
  switch (node.type) {
    case 'heading':      return renderHeading(node, baseUrl);
    case 'paragraph':    return renderInline(node.children, baseUrl) + '\n';
    case 'list':         return renderList(node, depth, baseUrl);
    case 'table':        return renderTable(node, baseUrl);
    case 'code':         return renderCode(node);
    case 'quote':        return renderQuote(node, baseUrl);
    case 'image':        return renderImage(node, baseUrl) + '\n';
    case 'figure':       return renderFigure(node, baseUrl);
    case 'thematicBreak':
    case 'pageBreak':    return '---\n';
    case 'section':      return renderBlocks(node.children, depth, baseUrl);
    case 'footnote':     return renderBlocks(node.children, depth, baseUrl);
    default:             return '';
  }
}

// M1 — headings
function renderHeading(node, baseUrl) {
  return '#'.repeat(node.level) + ' ' + renderInline(node.children, baseUrl) + '\n';
}

// M3 — lists
const { renderList } = makeListRenderer(renderInline, renderBlock);

// M4 — GFM tables
function renderTable(node, baseUrl) {
  const allRows = [...(node.head ?? []), ...node.rows];
  if (allRows.length === 0) return '';

  const colCount = Math.max(...allRows.map((r) => r.cells.length));

  const grid = allRows.map((row) =>
    Array.from({ length: colCount }, (_, i) =>
      renderCellText(row.cells[i]?.children ?? [], baseUrl).replace(/\|/g, '\\|'),
    ),
  );

  const header    = grid[0] ?? [];
  const separator = header.map(() => '---');
  const body      = grid.slice(1);

  const toRow = (cells) => '| ' + cells.join(' | ') + ' |';

  let out = toRow(header) + '\n' + toRow(separator) + '\n';
  for (const row of body) out += toRow(row) + '\n';
  return out;
}

// M5 — fenced code blocks
function renderCode(node) {
  const lang = node.language ?? '';
  return '```' + lang + '\n' + node.value + '\n```\n';
}

// M5 — blockquotes
function renderQuote(node, baseUrl) {
  const inner = renderBlocks(node.children, 0, baseUrl).trimEnd();
  return inner.split('\n').map((line) => '> ' + line).join('\n') + '\n';
}

// M6 — standalone image
// A data: URI src (megabytes of base64) has no place inline in Markdown —
// with a real link, use the standard image syntax; with no usable link,
// fall back to plain alt text, or a generic marker if there's neither.
function renderImage(node, baseUrl) {
  const src = displayImageSrc(node.src, baseUrl);
  if (src) return `![${node.alt ?? ''}](${src})`;
  if (node.alt) return node.alt;
  return '*(image)*';
}

// A table cell may hold multiple blocks or block content (e.g. a nested
// <table>) rather than one run of inline nodes. Blocks are joined with a raw
// `<br>` — valid inside a GFM table cell — since a real GFM table row can't.
function renderCellText(nodes, baseUrl) {
  const sep = nodes.length > 1 && nodes.every(isBlockNode) ? '<br>' : '';
  return nodes.map((n) => renderInlineNode(n, baseUrl)).join(sep);
}


// M6 — figure with optional caption
function renderFigure(node, baseUrl) {
  const img = renderImage(node.image, baseUrl);
  return node.caption ? `${img}\n*${node.caption}*\n` : `${img}\n`;
}

// ─── Inline rendering (M2) ───────────────────────────────────────────────────

function renderInline(nodes, baseUrl) {
  return nodes.map((n) => renderInlineNode(n, baseUrl)).join('');
}

function renderInlineNode(node, baseUrl) {
  switch (node.type) {
    case 'text':      return node.value;
    case 'strong':    return `**${renderInline(node.children, baseUrl)}**`;
    case 'emphasis':  return `*${renderInline(node.children, baseUrl)}*`;
    case 'codeSpan':  return `\`${node.value}\``;
    case 'styledSpan': return renderStyledSpan(node, baseUrl);
    case 'link': {
      const title = node.title ? ` "${node.title}"` : '';
      return `[${renderInline(node.children, baseUrl)}](${node.href}${title})`;
    }
    case 'lineBreak': return '  \n';
    case 'image':     return renderImage(node, baseUrl);
    // Falls through here when a table cell holds block content (e.g. a
    // nested <table>) instead of plain inline runs — flatten it sensibly
    // rather than silently dropping it.
    case 'paragraph':
    case 'heading':
    case 'quote':
    case 'section':   return renderInline(node.children, baseUrl);
    case 'code':      return '`' + node.value.replace(/\r?\n/g, ' ') + '`';
    case 'list':
      return node.items.map((item) => item.children.map((n) => renderInlineNode(n, baseUrl)).join(' ')).join('; ');
    case 'table':     return flattenNestedTable(node, baseUrl);
    default:          return '';
  }
}

// Markdown has no color/background/underline/size — those are dropped, but
// bold/italic (the two the format can express) are applied like <strong>/<em>.
function renderStyledSpan(node, baseUrl) {
  let inner = renderInline(node.children, baseUrl);
  if (node.style?.italic) inner = `*${inner}*`;
  if (node.style?.bold) inner = `**${inner}**`;
  return inner;
}

function flattenNestedTable(table, baseUrl) {
  return tableRows(table)
    .map((row) => row.cells.map((c) => renderCellText(c.children ?? [], baseUrl)).filter((t) => t.trim() !== ''))
    .filter((cells) => cells.length > 0)
    .map((cells) => cells.join(', '))
    .join('; ');
}
