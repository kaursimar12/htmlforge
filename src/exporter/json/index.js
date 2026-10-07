/**
 * JSON exporter — IDM → a plain JSON document.
 * Milestones: J1 (mirror IDM as {type,children}/{type,value}) ·
 *             J2 (tables → flat headers/rows) · J3 (meta + asset manifest)
 *
 * Top-level shape: { meta, content, assets } — not a wrapped "document" node.
 */

import { displayImageSrc } from '../../utils/html.js';
import { isBlockNode } from '../../idm/flatten-text.js';

const MIME_BY_EXTENSION = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  bmp: 'image/bmp',
};

export const jsonExporter = {
  id: 'json',
  extensions: ['.json'],
  mimeType: 'application/json',

  export(document, ctx) {
    const baseUrl = ctx?.options?.baseUrl;
    const envelope = {
      meta: document.metadata ?? {},
      content: mapJson(document.children, baseUrl),
      assets: collectAssets(document.children),
    };
    return Buffer.from(JSON.stringify(envelope, null, 2) + '\n', 'utf8');
  },
};

// ─── J1 — mirror the IDM ──────────────────────────────────────────────────────
// `baseUrl` (from convert()'s own options) is threaded through purely so it
// can reach an `image` node wherever one is nested — see nodeToJson's own
// 'image' case and flattenTextNode's identical one for table-cell content.

// Converts a list of IDM nodes to JSON, dropping any resulting table that
// turned out to carry zero content (see isEmptyTable) — e.g. a decorative
// spacer/divider <table> whose only cell is `&nbsp;`. This never drops a
// table with real (even blank) data cells, only ones that are 100% empty.
function mapJson(nodes, baseUrl) {
  return nodes.map((n) => nodeToJson(n, baseUrl)).filter((n) => !isEmptyTable(n));
}

function isEmptyTable(n) {
  return n.type === 'table' && n.headers.length === 0 && n.rows.length === 0;
}

function nodeToJson(node, baseUrl) {
  switch (node.type) {
    case 'section':
    case 'footnote':
    case 'quote':
    case 'strong':
    case 'emphasis':
      return { type: node.type, children: mapJson(node.children, baseUrl) };

    case 'heading': {
      const json = { type: 'heading', level: node.level, children: mapJson(node.children, baseUrl) };
      if (node.align) json.align = node.align;
      return json;
    }

    case 'paragraph': {
      const json = { type: 'paragraph', children: mapJson(node.children, baseUrl) };
      if (node.align) json.align = node.align;
      return json;
    }

    case 'styledSpan':
      return { type: 'styledSpan', style: node.style, children: mapJson(node.children, baseUrl) };

    case 'list':
      return {
        type: 'list',
        ordered: node.ordered,
        start: node.start,
        children: mapJson(node.items, baseUrl),
      };

    case 'listItem':
      return { type: 'listItem', children: mapJson(node.children, baseUrl) };

    case 'table':
      return tableToJson(node, baseUrl); // J2

    case 'code':
      return { type: 'code', language: node.language, value: node.value };

    case 'image':
      // A data: URI src is dropped here (megabytes of base64 has no place
      // in a readable JSON preview) — the real bytes are still available in
      // the top-level "assets" manifest for anything that needs them.
      return { type: 'image', src: displayImageSrc(node.src, baseUrl), alt: node.alt };

    case 'figure': {
      const children = [nodeToJson(node.image, baseUrl)];
      if (node.caption) children.push({ type: 'text', value: node.caption });
      return { type: 'figure', children };
    }

    case 'thematicBreak':
    case 'pageBreak':
    case 'lineBreak':
      return { type: node.type };

    case 'text':
    case 'codeSpan':
      return { type: node.type, value: node.value };

    case 'link':
      return {
        type: 'link',
        href: node.href,
        title: node.title,
        children: mapJson(node.children, baseUrl),
      };

    default:
      return { type: node.type };
  }
}

// ─── J2 — tables → flat headers/rows ──────────────────────────────────────────

function tableToJson(table, baseUrl) {
  const headRows = (table.head ?? []).map((r) => rowToValues(r, baseUrl)).filter(rowHasContent);
  const bodyRows = table.rows.map((r) => rowToValues(r, baseUrl)).filter(rowHasContent);

  const headers = headRows[0] ?? [];
  const rows = [...headRows.slice(1), ...bodyRows];

  const json = { type: 'table', headers, rows };
  if (table.caption) json.caption = table.caption;
  if (table.border) json.border = table.border;
  return json;
}

function rowToValues(row, baseUrl) {
  return row.cells.map((c) => cellValue(c, baseUrl));
}

// A row where every cell resolved to an empty string carries no information
// (e.g. a decorative <tr><td>&nbsp;</td></tr> divider row) — drop it. A row
// with even one real (possibly blank-looking but structurally present) value
// is left alone.
function rowHasContent(values) {
  return values.some((v) => v !== '');
}

// A cell with plain inline content (text, bold, links, ...) becomes a flat
// string, as before. A cell with block content (a nested <table>, multiple
// paragraphs, ...) is kept as real nested structure via nodeToJson instead of
// being collapsed into one string — JSON is the one output format in this
// project that can actually hold that; CSV/TXT/Markdown cannot and still flatten.
function cellValue(cell, baseUrl) {
  const nodes = cell.children ?? [];
  if (nodes.length === 0) return '';
  if (nodes.every((n) => !isBlockNode(n))) return flattenText(nodes, baseUrl);
  const converted = mapJson(nodes, baseUrl);
  if (converted.length === 0) return ''; // e.g. the cell held only a now-pruned empty table
  return converted.length === 1 ? converted[0] : converted;
}

function flattenText(nodes, baseUrl) {
  return nodes.map((n) => flattenTextNode(n, baseUrl)).join('');
}

function flattenTextNode(node, baseUrl) {
  switch (node.type) {
    case 'text':
      return node.value;
    case 'strong':
    case 'emphasis':
    case 'link':
    case 'styledSpan':
      return flattenText(node.children, baseUrl);
    case 'codeSpan':
      return node.value;
    case 'lineBreak':
      return '\n';
    case 'image':
      return node.alt ?? displayImageSrc(node.src, baseUrl) ?? '';
    default:
      return '';
  }
}

// ─── J3 — asset manifest ──────────────────────────────────────────────────────
// Scans the IDM for image references directly rather than reusing whatever
// `resolveAssets` produced — this exporter never embeds bytes (no format
// concept of "media" here), so the manifest only ever needs each image's own
// src/mimeType, not fetched data.

function collectAssets(nodes) {
  const images = [];
  walkForImages(nodes, images);
  return images.map((image, i) => ({
    id: `asset-${i + 1}`,
    src: image.src,
    mimeType: guessMimeType(image.src),
  }));
}

function walkForImages(nodes, out) {
  for (const node of nodes ?? []) {
    if (!node) continue;

    if (node.type === 'image') {
      out.push(node);
      continue;
    }
    if (node.type === 'figure') {
      out.push(node.image);
      continue;
    }
    if (node.type === 'table') {
      for (const row of [...(node.head ?? []), ...node.rows]) {
        for (const cell of row.cells) walkForImages(cell.children ?? [], out);
      }
      continue;
    }

    if (node.children) walkForImages(node.children, out);
    if (node.items) for (const item of node.items) walkForImages(item.children, out);
  }
}

function guessMimeType(src) {
  if (!src) return undefined;

  if (src.startsWith('data:')) {
    const match = /^data:([^;,]+)/.exec(src);
    return match ? match[1] : undefined;
  }

  const clean = src.split('?')[0].split('#')[0];
  const ext = clean.slice(clean.lastIndexOf('.') + 1).toLowerCase();
  return MIME_BY_EXTENSION[ext];
}
