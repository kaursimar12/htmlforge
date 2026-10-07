/**
 * Tree builder: turns the token stream into a DOM-like tree.
 *
 * Applies auto-close rules so that real-world HTML without explicit close tags
 * (e.g. `<p>text<p>next`) is handled correctly.
 */

import { tokenize } from './tokenizer.js';

/** Tags that auto-close when another tag of the same "group" opens. */
const AUTO_CLOSE = {
  p: new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'table', 'div', 'blockquote', 'pre', 'hr']),
  li: new Set(['li']),
  dt: new Set(['dt', 'dd']),
  dd: new Set(['dt', 'dd']),
  // A new <tr> implicitly closes a still-open <td>/<th> too, not just
  // another <td>/<th> — real-world malformed markup (missing </td>) is
  // exactly where this matters: `<tr><td>1<td>2<td>3<tr><td>4` needs row 2's
  // <tr> to close cell "3" before starting the new row, or "3" silently
  // swallows the entire next row as nested content instead of a sibling.
  td: new Set(['td', 'th', 'tr']),
  th: new Set(['td', 'th', 'tr']),
  tr: new Set(['tr']),
  option: new Set(['option', 'optgroup']),
  colgroup: new Set(['colgroup']),
  thead: new Set(['tbody', 'tfoot']),
  tbody: new Set(['thead', 'tbody', 'tfoot']),
  tfoot: new Set(['thead', 'tbody']),
};

export function buildTree(html) {
  const root = { type: 'element', name: '#root', attrs: {}, children: [], parent: null };
  const stack = [root];
  const tokens = tokenize(html);

  const top = () => stack[stack.length - 1];

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.type === 'text') {
      // A whitespace-only text token is only insignificant when BOTH
      // neighbors are block-level (or absent) — e.g. the indentation between
      // two sibling <p>s. When either neighbor is inline content (a text
      // run, or an inline element like <span>/<b>), the same whitespace is
      // the real, meaningful space between two words/runs on one line
      // (`<span>Bold</span> <span>Italic</span>`) and must be kept — dropping
      // it just because the *container* (e.g. <td>) happens to be block-level
      // glues adjacent inline runs together with no separator at all.
      const insignificant = token.value.trim() === '' && isBlockContext(top()) &&
        !isPrecededByInline(top().children) && !isFollowedByInline(tokens, i);
      if (insignificant) continue;
      const node = { type: 'text', value: token.value, parent: top() };
      top().children.push(node);
      continue;
    }

    if (token.type === 'comment') continue;

    if (token.type === 'close') {
      // Pop the stack until we find the matching open tag.
      for (let s = stack.length - 1; s >= 1; s--) {
        if (stack[s].name === token.name) {
          stack.length = s;
          break;
        }
      }
      continue;
    }

    // token.type === 'open'
    const { name, attrs, selfClosing } = token;

    // Auto-close open elements before inserting this one.
    autoClose(stack, name);

    const el = { type: 'element', name, attrs, children: [], parent: top() };
    top().children.push(el);
    if (!selfClosing) stack.push(el);
  }

  return root;
}

// Closing one level can reveal another that also needs to close for the same
// incoming tag — e.g. an incoming <tr> closes an open <td>, which reveals the
// <td>'s own open <tr> that must ALSO close for the same incoming <tr> (see
// the td/th AUTO_CLOSE comment above) — so this keeps popping the new top of
// stack, not just the original one, until a level doesn't match.
function autoClose(stack, incoming) {
  while (stack.length >= 2 && AUTO_CLOSE[stack[stack.length - 1].name]?.has(incoming)) {
    stack.length -= 1;
  }
}

const BLOCK_ELEMENTS = new Set([
  'div', 'p', 'ul', 'ol', 'li', 'table', 'tr', 'td', 'th',
  'thead', 'tbody', 'tfoot', 'blockquote', 'pre', 'h1', 'h2',
  'h3', 'h4', 'h5', 'h6', 'section', 'article', 'main',
  'header', 'footer', 'nav', 'figure', 'figcaption', 'hr',
]);

function isBlockContext(el) {
  return BLOCK_ELEMENTS.has(el.name) || el.name === '#root';
}

// The already-built previous sibling, if any: real text (even if itself
// whitespace — that's still "inline content" for this purpose) or an
// inline element (anything not in BLOCK_ELEMENTS) both count as inline.
function isPrecededByInline(children) {
  const prev = children[children.length - 1];
  if (!prev) return false;
  if (prev.type === 'text') return true;
  return !BLOCK_ELEMENTS.has(prev.name);
}

// Looks ahead in the raw token stream (skipping comments, which produce no
// node at all) for what comes right after this whitespace token: real text,
// or an opening tag for an inline element, both count as inline. A closing
// tag (nothing left in this container) or an opening tag for a block
// element means the block side, not the inline side.
function isFollowedByInline(tokens, index) {
  for (let j = index + 1; j < tokens.length; j++) {
    const t = tokens[j];
    if (t.type === 'comment') continue;
    if (t.type === 'text') return true;
    if (t.type === 'open') return !BLOCK_ELEMENTS.has(t.name);
    return false; // 'close'
  }
  return false;
}
