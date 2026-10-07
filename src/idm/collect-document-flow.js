/**
 * Walks the IDM in document order, keeping every block-level node — headings,
 * paragraphs, tables, lists, quotes, code blocks, images/figures — instead of
 * discarding everything but tables (contrast collect-tables.js, which throws
 * away headings/paragraphs once it's recorded a heading as a table's label).
 * Used by the XLSX exporter to lay out a whole document as one sheet, in the
 * same top-to-bottom order it appears in the HTML.
 *
 * Only recurses into `section` nodes' children (real `<section>`/`<article>`/
 * `<main>` wrappers) — a plain `<div>` never survives as its own IDM node
 * (dom-to-idm.js's `div` case flattens straight into its parent's children),
 * so by the time a node reaches here it's already one of the block types
 * below. List/quote/code/table/figure are leaf blocks for layout purposes:
 * their own internal structure (list items, table rows, a figure's image +
 * caption) is rendered by the exporter, not flattened further into the
 * document-order sequence.
 */

const BLOCK_TYPES = new Set(['heading', 'paragraph', 'table', 'list', 'quote', 'code', 'image', 'figure']);

export function collectDocumentFlow(nodes, out = []) {
  for (const node of nodes ?? []) {
    if (!node) continue;

    if (node.type === 'section') {
      collectDocumentFlow(node.children, out);
      continue;
    }

    if (BLOCK_TYPES.has(node.type)) {
      out.push(node);
      continue;
    }

    // Anything else (thematicBreak, pageBreak, svgShapes, ...) has no XLSX
    // row representation yet — silently skipped, same policy collect-tables.js
    // already applies to non-table/non-heading nodes.
  }
  return out;
}
