/**
 * Flatten IDM inline/block content down to plain text, no markup syntax —
 * shared by any exporter/heuristic that needs a single string representation
 * of a node's content (CSV's cell values, DOCX's content-based column-width
 * heuristic, XLSX's cell values). Originally private to the CSV exporter;
 * pulled out here once a second consumer needed the exact same logic.
 */

import { displayImageSrc } from '../utils/html.js';

const BLOCK_NODE_TYPES = new Set(['paragraph', 'heading', 'quote', 'section', 'list', 'table', 'code']);

export function isBlockNode(node) {
  return BLOCK_NODE_TYPES.has(node.type);
}

/** A table's header + body rows, in document order — the "every row,
 * regardless of which part of the table it came from" concat every
 * row-by-row table walk needs (CSV/TXT/Markdown's rendering, this file's own
 * flattenNestedTable below). */
export function tableRows(table) {
  return [...(table.head ?? []), ...table.rows];
}

/** A table cell's content (inline or block) as one plain-text string. */
export function flattenCellText(cell) {
  const nodes = cell.children ?? [];
  const sep = nodes.length > 1 && nodes.every(isBlockNode) ? '\n' : '';
  return nodes.map(flattenNode).join(sep);
}

/** Any array of IDM inline nodes, concatenated with no separator. */
export function flattenInline(nodes) {
  return nodes.map(flattenNode).join('');
}

export function flattenNode(node) {
  switch (node.type) {
    case 'text':
      return node.value;
    case 'strong':
    case 'emphasis':
    case 'link':
    case 'styledSpan':
      return flattenInline(node.children);
    case 'codeSpan':
      return node.value;
    case 'lineBreak':
      return '\n';
    case 'image':
      return node.alt ?? displayImageSrc(node.src) ?? '';
    // Reached when a table cell holds block content (e.g. a nested <table>)
    // instead of plain inline runs — flatten it to plain text too.
    case 'paragraph':
    case 'heading':
    case 'quote':
    case 'section':
      return flattenInline(node.children);
    case 'code':
      return node.value;
    case 'list':
      return node.items.map((item) => item.children.map(flattenNode).join(' ')).join('; ');
    case 'table':
      return flattenNestedTable(node);
    default:
      return '';
  }
}

function flattenNestedTable(table) {
  return tableRows(table)
    .map((row) => row.cells.map(flattenCellText).filter((t) => t.trim() !== ''))
    .filter((cells) => cells.length > 0)
    .map((cells) => cells.join(' | '))
    .join('\n');
}
