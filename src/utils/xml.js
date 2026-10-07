/**
 * Minimal dependency-free XML parser.
 *
 * Targets the well-formed XML inside OOXML packages (DOCX/XLSX/PPTX).
 * Namespace prefixes are preserved verbatim (e.g. `w:p`, `a:t`); traversal
 * helpers match on either the full prefixed name or the local part.
 */

const NAMED_ENTITIES = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'",
};

export function decodeEntities(value) {
  if (!value.includes('&')) return value;
  return value.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (match, body) => {
    if (body[0] === '#') {
      const code =
        body[1] === 'x' || body[1] === 'X'
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      return Number.isFinite(code) ? safeFromCodePoint(code) : match;
    }
    return NAMED_ENTITIES[body] ?? match;
  });
}

function safeFromCodePoint(code) {
  try { return String.fromCodePoint(code); } catch { return ''; }
}

export function parseXml(input) {
  const root = { type: 'element', name: '#root', attrs: {}, children: [] };
  const stack = [root];
  let i = 0;
  const len = input.length;
  const top = () => stack[stack.length - 1];

  while (i < len) {
    const lt = input.indexOf('<', i);
    if (lt === -1) { pushText(top(), input.slice(i)); break; }
    if (lt > i) pushText(top(), input.slice(i, lt));

    if (input.startsWith('<!--', lt)) {
      const end = input.indexOf('-->', lt + 4);
      i = end === -1 ? len : end + 3;
      continue;
    }
    if (input.startsWith('<![CDATA[', lt)) {
      const end = input.indexOf(']]>', lt + 9);
      const data = input.slice(lt + 9, end === -1 ? len : end);
      top().children.push({ type: 'text', value: data });
      i = end === -1 ? len : end + 3;
      continue;
    }
    if (input.startsWith('<?', lt) || input.startsWith('<!', lt)) {
      const end = input.indexOf('>', lt + 1);
      i = end === -1 ? len : end + 1;
      continue;
    }
    if (input[lt + 1] === '/') {
      const end = input.indexOf('>', lt + 2);
      const name = input.slice(lt + 2, end === -1 ? len : end).trim();
      for (let s = stack.length - 1; s >= 1; s--) {
        if (stack[s].name === name) { stack.length = s; break; }
      }
      i = end === -1 ? len : end + 1;
      continue;
    }

    const end = findTagEnd(input, lt);
    if (end === -1) { pushText(top(), input.slice(lt)); break; }
    const raw = input.slice(lt + 1, end);
    const selfClosing = raw.endsWith('/');
    const body = selfClosing ? raw.slice(0, -1) : raw;
    const { name, attrs } = parseTag(body);
    const element = { type: 'element', name, attrs, children: [] };
    top().children.push(element);
    if (!selfClosing) stack.push(element);
    i = end + 1;
  }

  const elementChildren = root.children.filter((c) => c.type === 'element');
  return elementChildren.length === 1 ? elementChildren[0] : root;
}

function findTagEnd(input, lt) {
  let quote = null;
  for (let j = lt + 1; j < input.length; j++) {
    const ch = input[j];
    if (quote) { if (ch === quote) quote = null; }
    else if (ch === '"' || ch === "'") { quote = ch; }
    else if (ch === '>') return j;
  }
  return -1;
}

function parseTag(body) {
  const trimmed = body.trim();
  let k = 0;
  while (k < trimmed.length && !isSpace(trimmed[k])) k++;
  const name = trimmed.slice(0, k);
  const attrs = {};
  const attrRe = /([^\s=/]+)\s*(?:=\s*("([^"]*)"|'([^']*)'|([^\s/]+)))?/g;
  attrRe.lastIndex = k;
  let m;
  while ((m = attrRe.exec(trimmed)) !== null) {
    if (m[0].trim() === '') break;
    attrs[m[1]] = decodeEntities(m[3] ?? m[4] ?? m[5] ?? '');
  }
  return { name, attrs };
}

function isSpace(ch) {
  return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r';
}

function pushText(parent, raw) {
  if (raw === '') return;
  parent.children.push({ type: 'text', value: decodeEntities(raw) });
}

function nameMatches(el, query) {
  if (el.name === query) return true;
  const colon = el.name.indexOf(':');
  return colon !== -1 && el.name.slice(colon + 1) === query;
}

export function childElements(el, name) {
  const out = [];
  for (const child of el.children) {
    if (child.type === 'element' && (name === undefined || nameMatches(child, name)))
      out.push(child);
  }
  return out;
}

export function firstChild(el, name) {
  for (const child of el.children) {
    if (child.type === 'element' && nameMatches(child, name)) return child;
  }
  return undefined;
}

export function find(el, name) {
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    if (nameMatches(child, name)) return child;
    const nested = find(child, name);
    if (nested) return nested;
  }
  return undefined;
}

export function findAll(el, name) {
  const out = [];
  const walk = (node) => {
    for (const child of node.children) {
      if (child.type !== 'element') continue;
      if (nameMatches(child, name)) out.push(child);
      walk(child);
    }
  };
  walk(el);
  return out;
}

export function textContent(node) {
  if (node.type === 'text') return node.value;
  let out = '';
  for (const child of node.children) out += textContent(child);
  return out;
}

export function attr(el, name) {
  if (name in el.attrs) return el.attrs[name];
  for (const key of Object.keys(el.attrs)) {
    const colon = key.indexOf(':');
    if (colon !== -1 && key.slice(colon + 1) === name) return el.attrs[key];
  }
  return undefined;
}
