/**
 * Normalizer: tidies the IDM before it reaches an exporter.
 *
 * Collapses insignificant whitespace and drops empty nodes that would produce
 * empty elements in the output.
 */

export function normalizeDocument(document) {
  return { ...document, children: normalizeBlocks(document.children) };
}

function normalizeBlocks(nodes) {
  const out = [];
  for (const node of nodes) {
    const normalized = normalizeBlock(node);
    if (normalized) out.push(normalized);
  }
  return out;
}

function normalizeBlock(node) {
  switch (node.type) {
    case 'section': {
      const children = normalizeBlocks(node.children);
      if (children.length === 0) return null;
      return { ...node, children };
    }
    case 'quote': {
      const children = normalizeBlocks(node.children);
      if (children.length === 0) return null;
      return { ...node, children };
    }
    case 'paragraph': {
      const children = normalizeInline(node.children);
      if (children.length === 0) return null;
      return { ...node, children };
    }
    case 'heading': {
      const children = normalizeInline(node.children);
      if (children.length === 0) return null;
      return { ...node, children };
    }
    case 'list': {
      const items = node.items
        .map(normalizeListItem)
        .filter((item) => item !== null);
      if (items.length === 0) return null;
      return { ...node, items };
    }
    case 'table': {
      const head = node.head ? normalizeTableRows(node.head) : undefined;
      const rows = normalizeTableRows(node.rows);
      return { ...node, head, rows };
    }
    default:
      return node;
  }
}

function normalizeTableRows(rows) {
  return rows.map((row) => ({
    ...row,
    cells: row.cells.map((cell) => ({ ...cell, children: normalizeCellChildren(cell.children) })),
  }));
}

// A cell's children are either plain inline nodes (the common case) or block
// nodes (when the source <td> held flow content like a nested table) — pick
// the matching normalizer instead of assuming one shape.
const INLINE_NODE_TYPES = new Set(['text', 'strong', 'emphasis', 'codeSpan', 'link', 'lineBreak', 'styledSpan']);

function normalizeCellChildren(children) {
  const looksInline = children.every((n) => INLINE_NODE_TYPES.has(n.type) || n.type === 'image');
  return looksInline ? normalizeInline(children) : normalizeBlocks(children);
}

function normalizeListItem(item) {
  const children = normalizeBlocks(item.children);
  if (children.length === 0) return null;
  return { ...item, children };
}

function normalizeInline(nodes) {
  const out = [];
  for (const node of nodes) {
    if (node.type === 'text') {
      const value = node.value.replace(/\s+/g, ' ');
      if (value === '') continue;
      out.push({ type: 'text', value });
    } else if (node.type === 'strong' || node.type === 'emphasis' || node.type === 'link' || node.type === 'styledSpan') {
      const children = normalizeInline(node.children);
      if (children.length === 0) continue;
      out.push({ ...node, children });
    } else {
      out.push(node);
    }
  }
  return trimEdges(out);
}

function trimEdges(nodes) {
  const result = [...nodes];
  const first = result[0];
  if (first?.type === 'text') {
    const value = first.value.replace(/^\s+/, '');
    if (value === '') result.shift();
    else result[0] = { type: 'text', value };
  }
  const last = result[result.length - 1];
  if (last?.type === 'text') {
    const value = last.value.replace(/\s+$/, '');
    if (value === '') result.pop();
    else result[result.length - 1] = { type: 'text', value };
  }
  return result;
}
