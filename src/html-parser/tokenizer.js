/**
 * Forgiving HTML tokenizer.
 *
 * Emits a stream of tokens from an HTML string. Handles void elements, raw-text
 * elements (script/style), self-closing tags, and attribute quoting styles.
 * Does NOT implement the full WHATWG spec — real-world editor output (Notion,
 * Quill, TipTap, CMS) is well-structured enough for this pragmatic subset.
 */

/** Elements that never have children. */
const VOID_ELEMENTS = new Set([
  'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
  'link', 'meta', 'param', 'source', 'track', 'wbr',
]);

/** Elements whose content is raw text (not parsed as HTML). */
const RAW_TEXT_ELEMENTS = new Set(['script', 'style', 'noscript', 'textarea', 'template']);

export function tokenize(html) {
  const tokens = [];
  let i = 0;
  const len = html.length;

  while (i < len) {
    const lt = html.indexOf('<', i);

    if (lt === -1) {
      // Remaining text after the last tag.
      const text = html.slice(i);
      if (text) tokens.push({ type: 'text', value: decodeEntities(text) });
      break;
    }

    // Text content before this '<'.
    if (lt > i) {
      tokens.push({ type: 'text', value: decodeEntities(html.slice(i, lt)) });
    }

    // Comment <!-- ... -->
    if (html.startsWith('<!--', lt)) {
      const end = html.indexOf('-->', lt + 4);
      i = end === -1 ? len : end + 3;
      tokens.push({ type: 'comment' });
      continue;
    }

    // Declaration / PI — skip.
    if (html.startsWith('<?', lt) || html.startsWith('<!', lt)) {
      const end = html.indexOf('>', lt + 1);
      i = end === -1 ? len : end + 1;
      continue;
    }

    // Closing tag </name>
    if (html[lt + 1] === '/') {
      const end = html.indexOf('>', lt + 2);
      const name = html.slice(lt + 2, end === -1 ? len : end).trim().toLowerCase();
      tokens.push({ type: 'close', name });
      i = end === -1 ? len : end + 1;
      continue;
    }

    // Opening tag
    const end = findTagEnd(html, lt);
    if (end === -1) {
      tokens.push({ type: 'text', value: html.slice(lt) });
      break;
    }

    const raw = html.slice(lt + 1, end);
    const selfClosing = raw.endsWith('/');
    const body = selfClosing ? raw.slice(0, -1) : raw;
    const { name, attrs } = parseOpenTag(body);

    const isVoid = VOID_ELEMENTS.has(name);
    tokens.push({ type: 'open', name, attrs, selfClosing: selfClosing || isVoid });
    i = end + 1;

    // For raw-text elements, swallow everything until the matching close tag.
    if (RAW_TEXT_ELEMENTS.has(name) && !isVoid) {
      const closeTag = `</${name}`;
      const closeIdx = html.toLowerCase().indexOf(closeTag, i);
      if (closeIdx !== -1) {
        const rawText = html.slice(i, closeIdx);
        if (rawText) tokens.push({ type: 'text', value: rawText });
        const afterClose = html.indexOf('>', closeIdx + closeTag.length);
        i = afterClose === -1 ? len : afterClose + 1;
        tokens.push({ type: 'close', name });
      } else {
        i = len;
      }
    }
  }

  return tokens;
}

function findTagEnd(html, lt) {
  let quote = null;
  for (let j = lt + 1; j < html.length; j++) {
    const ch = html[j];
    if (quote) { if (ch === quote) quote = null; }
    else if (ch === '"' || ch === "'") { quote = ch; }
    else if (ch === '>') return j;
  }
  return -1;
}

function parseOpenTag(body) {
  const trimmed = body.trim();
  let k = 0;
  while (k < trimmed.length && !/[\s>]/.test(trimmed[k])) k++;
  const name = trimmed.slice(0, k).toLowerCase();
  const attrs = {};

  // Match: name="val" | name='val' | name=val | name (boolean)
  const attrRe = /([^\s=/"']+)\s*(?:=\s*(?:"([^"]*)"|'([^']*)'|([^\s>"']*)))?/g;
  attrRe.lastIndex = k;
  let m;
  while ((m = attrRe.exec(trimmed)) !== null) {
    if (!m[0].trim()) break;
    const attrName = m[1].toLowerCase();
    const attrVal = m[2] ?? m[3] ?? m[4] ?? '';
    attrs[attrName] = decodeEntities(attrVal);
  }
  return { name, attrs };
}

const HTML_ENTITIES = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
  mdash: '—', ndash: '–', lsquo: '‘', rsquo: '’',
  ldquo: '“', rdquo: '”', hellip: '…', copy: '©',
  reg: '®', trade: '™', euro: '€',
};

function decodeEntities(value) {
  if (!value.includes('&')) return value;
  return value.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (match, body) => {
    if (body[0] === '#') {
      const code =
        body[1] === 'x' || body[1] === 'X'
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      if (Number.isFinite(code)) {
        try { return String.fromCodePoint(code); } catch { return match; }
      }
      return match;
    }
    return HTML_ENTITIES[body] ?? match;
  });
}
