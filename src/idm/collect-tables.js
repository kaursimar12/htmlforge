/**
 * Walks the IDM in document order, pairing each `TableNode` with its own
 * `<caption>` or (failing that) the nearest preceding heading text — the
 * "find every table, labeled" walk shared by every exporter that treats each
 * table as its own independent unit (CSV's concatenated sections, XLSX's one
 * sheet per table). Deliberately does not recurse into `tableCell.children`
 * — a table nested inside another table's cell is flattened to text there
 * (see flatten-text.js's `flattenNestedTable`), not promoted to its own
 * top-level entry.
 */

import { flattenInline } from './flatten-text.js';

export function collectTables(nodes, state = { lastHeading: undefined }, out = []) {
  for (const node of nodes ?? []) {
    if (!node) continue;

    if (node.type === 'heading') {
      state.lastHeading = flattenInline(node.children ?? []).trim() || undefined;
      continue;
    }

    if (node.type === 'table') {
      out.push({ table: node, label: node.caption ?? state.lastHeading });
      state.lastHeading = undefined;
      continue;
    }

    if (node.children) collectTables(node.children, state, out);
    if (node.items) for (const item of node.items) collectTables(item.children, state, out);
  }
  return out;
}
