/**
 * A small, purpose-built PDF reader for asserting on this project's own
 * exporter output — not a general-purpose PDF parser (same spirit as
 * tests/helpers/docx.js's zip+XML parsing: verify real structure, never
 * regex raw bytes, but only the subset of PDF syntax this exporter itself
 * produces).
 */
import { inflateSync } from 'node:zlib';

function skipWhitespace(text, pos) {
  while (pos < text.length && /[\s]/.test(text[pos])) pos++;
  // PDF comments (`%...` to end of line) are also whitespace-equivalent
  // between tokens, but this project's own writer never emits one mid-object
  // (only in the header), so no comment-skipping is needed here.
  return pos;
}

function parseName(text, pos) {
  let end = pos + 1; // skip leading '/'
  while (end < text.length && !/[\s/()<>\[\]]/.test(text[end])) end++;
  return { value: text.slice(pos + 1, end), endPos: end };
}

// A hex string (`<...>`, e.g. an embedded font's `showHexText` glyph-ID
// operand) — distinct from a `<<...>>` dict, which parseValue/parseDict
// already check for before ever reaching this. Kept as the raw hex digit
// string rather than decoded: unlike parseString's decoded text, hex-string
// operands in this project's own content streams are glyph IDs, not
// characters, so there's no single meaningful decoded form to return.
function parseHexString(text, pos) {
  let end = pos + 1;
  while (end < text.length && text[end] !== '>') end++;
  return { value: text.slice(pos + 1, end), endPos: end + 1 };
}

function parseString(text, pos) {
  let end = pos + 1;
  let depth = 1;
  let out = '';
  while (end < text.length && depth > 0) {
    const ch = text[end];
    if (ch === '\\') { out += text[end + 1]; end += 2; continue; }
    if (ch === '(') depth++;
    if (ch === ')') { depth--; if (depth === 0) { end++; break; } }
    out += ch;
    end++;
  }
  return { value: out, endPos: end };
}

function parseNumber(text, pos) {
  let end = pos;
  if (text[end] === '+' || text[end] === '-') end++;
  while (end < text.length && /[0-9.]/.test(text[end])) end++;
  return { value: parseFloat(text.slice(pos, end)), endPos: end };
}

// A bare integer that's followed by another integer and then a literal `R`
// is an indirect reference (`N G R`), not two plain numbers — used inside
// dict/array values (`/Parent 2 0 R`), never inside a content stream (which
// never contains references at all).
function parseNumberOrRef(text, pos) {
  const first = parseNumber(text, pos);
  let p = skipWhitespace(text, first.endPos);
  if (!/[0-9]/.test(text[p] ?? '')) return first;
  const gen = parseNumber(text, p);
  let p2 = skipWhitespace(text, gen.endPos);
  if (text[p2] === 'R') return { value: { ref: first.value }, endPos: p2 + 1 };
  return first;
}

function parseArray(text, pos) {
  let p = skipWhitespace(text, pos + 1);
  const items = [];
  while (text[p] !== ']') {
    const { value, endPos } = parseValue(text, p);
    items.push(value);
    p = skipWhitespace(text, endPos);
  }
  return { value: items, endPos: p + 1 };
}

function parseDict(text, pos) {
  let p = skipWhitespace(text, pos + 2);
  const dict = {};
  while (!text.startsWith('>>', p)) {
    const { value: key, endPos: keyEnd } = parseName(text, p);
    p = skipWhitespace(text, keyEnd);
    const { value, endPos } = parseValue(text, p);
    dict[key] = value;
    p = skipWhitespace(text, endPos);
  }
  return { value: dict, endPos: p + 2 };
}

function parseValue(text, pos) {
  const p = skipWhitespace(text, pos);
  const ch = text[p];
  if (text.startsWith('<<', p)) return parseDict(text, p);
  if (ch === '[') return parseArray(text, p);
  if (ch === '/') return parseName(text, p);
  if (ch === '(') return parseString(text, p);
  if (ch === '<') return parseHexString(text, p);
  if (ch === '+' || ch === '-' || /[0-9]/.test(ch)) return parseNumberOrRef(text, p);
  if (text.startsWith('true', p)) return { value: true, endPos: p + 4 };
  if (text.startsWith('false', p)) return { value: false, endPos: p + 5 };
  if (text.startsWith('null', p)) return { value: null, endPos: p + 4 };
  throw new Error(`Unexpected PDF token at ${p}: ${JSON.stringify(text.slice(p, p + 30))}`);
}

/** Parses the classic xref table + trailer, and provides indirect-object resolution. */
export function parsePdf(buffer) {
  const text = buffer.toString('latin1');
  if (!text.startsWith('%PDF-')) throw new Error('Not a PDF: missing %PDF- header');

  const startxrefKeywordIdx = text.lastIndexOf('startxref');
  const xrefOffset = parseInt(text.slice(startxrefKeywordIdx).match(/startxref\s+(\d+)/)[1], 10);

  const xrefSection = text.slice(xrefOffset);
  if (!xrefSection.startsWith('xref')) throw new Error(`No xref table at recorded startxref offset ${xrefOffset}`);
  const header = xrefSection.match(/xref\s*\n0 (\d+)\s*\n/);
  const count = parseInt(header[1], 10);
  let entriesPos = header.index + header[0].length;
  const offsetByObjNum = new Map();
  for (let n = 0; n < count; n++) {
    const entry = xrefSection.slice(entriesPos, entriesPos + 20);
    const [offsetStr, , type] = entry.trim().split(/\s+/);
    if (type === 'n') offsetByObjNum.set(n, parseInt(offsetStr, 10));
    entriesPos += 20;
  }

  const trailerIdx = xrefSection.indexOf('trailer');
  const { value: trailer } = parseDict(xrefSection, skipWhitespace(xrefSection, trailerIdx + 'trailer'.length));

  return { text, offsetByObjNum, trailer };
}

/** Resolves an indirect reference (`{ref: N}`) or returns the value as-is if it already isn't one. */
export function resolveRef(pdf, value) {
  if (!value || typeof value !== 'object' || !('ref' in value)) return value;
  const offset = pdf.offsetByObjNum.get(value.ref);
  if (offset === undefined) throw new Error(`No xref entry for object ${value.ref}`);
  const header = pdf.text.slice(offset, offset + 40).match(/^(\d+)\s+(\d+)\s+obj\s*/);
  if (!header) throw new Error(`No "N G obj" header at offset ${offset} for object ${value.ref}`);
  const bodyStart = offset + header[0].length;
  const { value: dict, endPos } = parseValue(pdf.text, bodyStart);

  const afterDict = pdf.text.slice(endPos, endPos + 20);
  const streamMatch = afterDict.match(/^\s*stream\r?\n/);
  if (!streamMatch) return dict;

  const streamStart = endPos + streamMatch[0].length;
  const length = resolveRef(pdf, dict.Length);
  const rawStream = Buffer.from(pdf.text.slice(streamStart, streamStart + length), 'latin1');
  return { ...dict, __stream: rawStream };
}

export function getCatalog(pdf) {
  return resolveRef(pdf, pdf.trailer.Root);
}

export function getPages(pdf) {
  const catalog = getCatalog(pdf);
  const pagesRoot = resolveRef(pdf, catalog.Pages);
  return pagesRoot.Kids.map((kidRef) => resolveRef(pdf, kidRef));
}

/** Resolves a page's `/Contents` stream (inflating FlateDecode) and returns the raw operator text. */
export function getPageContentText(pdf, pageDict) {
  const contentObj = resolveRef(pdf, pageDict.Contents);
  return inflateSync(contentObj.__stream).toString('latin1');
}

/** Tokenizes content-stream operator text into `[{op, operands}]`. Content
 * streams never contain indirect references, so plain `parseNumber` (not
 * the ref-aware variant) is used here. */
export function tokenizeContentOps(text) {
  const ops = [];
  let operands = [];
  let pos = 0;
  while (pos < text.length) {
    pos = skipWhitespace(text, pos);
    if (pos >= text.length) break;
    const ch = text[pos];
    if (ch === '/') { const r = parseName(text, pos); operands.push(r.value); pos = r.endPos; continue; }
    if (ch === '(') { const r = parseString(text, pos); operands.push(r.value); pos = r.endPos; continue; }
    if (ch === '<') { const r = parseHexString(text, pos); operands.push(r.value); pos = r.endPos; continue; }
    if (ch === '[') { const r = parseArray(text, pos); operands.push(r.value); pos = r.endPos; continue; }
    if (ch === '+' || ch === '-' || /[0-9.]/.test(ch)) { const r = parseNumber(text, pos); operands.push(r.value); pos = r.endPos; continue; }
    let end = pos;
    while (end < text.length && !/[\s/(\[<]/.test(text[end])) end++;
    const word = text.slice(pos, end);
    pos = end;
    if (word === 'true' || word === 'false') { operands.push(word === 'true'); continue; }
    ops.push({ op: word, operands });
    operands = [];
  }
  return ops;
}

export function getPageContentOps(pdf, pageDict) {
  return tokenizeContentOps(getPageContentText(pdf, pageDict));
}

export function getAnnotations(pdf, pageDict) {
  return (pageDict.Annots ?? []).map((r) => resolveRef(pdf, r));
}

export function resolveResource(pdf, pageDict, category, resourceName) {
  const resources = resolveRef(pdf, pageDict.Resources);
  const categoryDict = resolveRef(pdf, resources[category]);
  return resolveRef(pdf, categoryDict?.[resourceName]);
}
