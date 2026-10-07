/**
 * Ordered/unordered list rendering (indent + marker + recurse into nested
 * lists/blocks) — identical algorithm shared by TXT and Markdown, differing
 * only in how each format renders inline content and other block types.
 * Those two behaviors are injected so each caller's own `renderInline`/
 * `renderBlock` (which may itself need to call back into `renderList` for a
 * nested list, or `renderListItem` — see each exporter's `renderBlock` 'list'
 * case) keeps working exactly as before, just sourced from one place.
 */
export function makeListRenderer(renderInline, renderBlock) {
  // `extra` is an opaque trailing value (e.g. Markdown/TXT's own `baseUrl`)
  // forwarded verbatim to every `renderInline`/`renderBlock` call so an
  // `image` node nested inside a list item still receives it — this module
  // has no opinion on what `extra` means, only that it must reach the leaf
  // renderers unchanged.
  function renderList(node, depth, extra) {
    return node.items.map((item, i) => renderListItem(item, node, i, depth, extra)).join('');
  }

  function renderListItem(item, list, index, depth, extra) {
    const indent = '  '.repeat(depth);
    const marker = list.ordered ? `${index + (list.start ?? 1)}.` : '-';

    const [first, ...rest] = item.children;
    let out = `${indent}${marker} `;

    if (first?.type === 'paragraph') {
      out += renderInline(first.children, extra) + '\n';
    } else if (first) {
      out += '\n' + renderBlock(first, depth + 1, extra);
    } else {
      out += '\n';
    }

    for (const child of rest) {
      if (child.type === 'list') {
        out += renderList(child, depth + 1, extra);
      } else {
        out += renderBlock(child, depth + 1, extra);
      }
    }

    return out;
  }

  return { renderList, renderListItem };
}
