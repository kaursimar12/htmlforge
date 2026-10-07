/**
 * Asset resolver (H6) — the single place in this project that knows *how* to
 * obtain an image's real bytes. Walks the IDM tree, collects every
 * ImageNode, resolves each distinct src to a normalized Asset record `{ id,
 * src, data, mimeType, extension, width, height }`, and stamps
 * ImageNode.assetId so exporters can look theirs up. Every exporter —
 * DOCX/PPTX/PDF/XLSX (binary, embed real bytes) and Markdown/TXT/JSON (text,
 * display-only) alike — consumes this same resolved shape; none of them
 * ever call `fetch`/`readFile` themselves (see convert()'s own pipeline in
 * src/index.js, which calls this exactly once per conversion and hands the
 * result to whichever exporter is running).
 *
 * Three src kinds are supported (classification/relative-path resolution
 * itself lives in utils/image-src.js — shared with the text exporters, which
 * need the identical resolution but must never actually load bytes):
 *   data:image/png;base64,...   → base64 (or percent-encoded text) decode
 *   http:// / https://          → fetched over the network (Node's built-in
 *                                  global fetch — no npm dependency)
 *   anything else                → a local filesystem path, read via
 *                                  node:fs, resolved against `options.baseUrl`
 *                                  when relative (see resolveImageSrc)
 *
 * Distinct sources are deduplicated up front (the `seen` map below) and then
 * resolved concurrently — since this function runs exactly once per
 * conversion, that dedup map *is* the per-conversion cache the same image
 * used three times in one document only ever gets fetched/read once.
 */

import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { detectFormat } from '../utils/magic.js';
import { classifyImageSrc, resolveImageSrc } from '../utils/image-src.js';
import { readImageDimensions } from '../utils/image-dimensions.js';

const DEFAULT_IMAGE_TIMEOUT_MS = 10_000;

// The one place mimeType maps back to a filename extension — every exporter
// that writes an embedded image to its own media/ folder (DOCX/PPTX/XLSX)
// reads `asset.extension` instead of maintaining its own copy of this map.
const EXTENSION_BY_MIME = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
};

// ─── Public API ───────────────────────────────────────────────────────────────

export async function resolveAssets(document, options = {}) {
  const { embedImages = true, strictAssets = false } = options;
  const assets = [];
  const warnings = [];
  const seen = new Map(); // src → assetId (dedup — the per-conversion cache)

  const toResolve = [];
  for (const img of collectImages(document)) {
    const src = img.src;
    if (!src) continue;

    // Same src already processed — reuse the same asset id
    if (seen.has(src)) {
      img.assetId = seen.get(src);
      continue;
    }

    const id = makeId(src);
    seen.set(src, id);
    img.assetId = id;
    toResolve.push({ id, src });
  }

  if (!embedImages) {
    for (const { id, src } of toResolve) {
      const mimeType = guessMimeFromSrc(src);
      assets.push({ id, type: 'image', mimeType, extension: EXTENSION_BY_MIME[mimeType] ?? 'png', src });
    }
    return { assets, warnings };
  }

  // Distinct sources resolve concurrently — real I/O (network fetch, disk
  // read) now, not a sync base64 decode, so serializing them would pay N
  // round-trips back-to-back for no reason.
  const resolved = await Promise.all(toResolve.map(({ id, src }) => resolveOne(id, src, options)));
  for (const asset of resolved) {
    if (asset.error) {
      warnings.push({
        code: 'MISSING_IMAGE',
        message: `Cannot resolve image "${asset.src}": ${asset.error}`,
        severity: 'warning',
      });
      if (strictAssets) throw new Error(`strictAssets: cannot resolve "${asset.src}": ${asset.error}`);
      const mimeType = guessMimeFromSrc(asset.src);
      assets.push({ id: asset.id, type: 'image', mimeType, extension: EXTENSION_BY_MIME[mimeType] ?? 'png', src: asset.src });
    } else {
      assets.push(asset);
    }
  }

  return { assets, warnings };
}

async function resolveOne(id, src, options) {
  try {
    const data = await fetchBytes(src, options);
    const detected = detectFormat(data);
    const mimeType = detected?.mimeType?.startsWith('image/') ? detected.mimeType : guessMimeFromSrc(src);
    // Prefer the extension `detectFormat` itself derived from the real magic
    // bytes (accurate even when the src's own file extension lies/is
    // missing — e.g. a data: URI, or a CDN URL with no extension at all);
    // only fall back to the mimeType→extension map when detection didn't
    // pan out (an SVG, or any format detectFormat doesn't recognize).
    const extension = detected?.mimeType === mimeType && detected.extension
      ? detected.extension.replace(/^\./, '')
      : (EXTENSION_BY_MIME[mimeType] ?? 'png');
    const dimensions = readImageDimensions(data, mimeType);
    return { id, type: 'image', src, data, mimeType, extension, width: dimensions?.width, height: dimensions?.height };
  } catch (err) {
    return { id, src, error: err instanceof Error ? err.message : String(err) };
  }
}

async function fetchBytes(src, options) {
  if (classifyImageSrc(src) === 'data') return decodeDataUri(src);

  // With no explicit baseUrl, a local path still needs SOME concrete root to
  // actually open a real file against — process.cwd(), the same implicit
  // root a plain `readFile(src)` would have used. resolveImageSrc itself
  // stays a pure "resolve against a base, or pass through if there's none"
  // function (see utils/image-src.js) — this cwd fallback is a resolver-only
  // policy, not something the text exporters' display-only calls should
  // ever apply.
  const resolved = resolveImageSrc(src, options.baseUrl ?? process.cwd());
  if (classifyImageSrc(resolved) === 'http') return fetchHttpImage(resolved, options);
  return readFile(resolved);
}

// ─── IDM walker ───────────────────────────────────────────────────────────────

// A generic recursive walk rather than an enumerated per-type switch — the
// enumerated version previously missed images nested directly in a table
// cell (only strong/emphasis/link wrapping wired), and any newly-added inline
// wrapper type (e.g. styledSpan) needs no further change here since it's
// picked up by the generic `.children` fallback.
function collectImages(document) {
  const out = [];
  walk(document.children, out);
  return out;
}

function walk(nodes, out) {
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
        for (const cell of row.cells) walk(cell.children ?? [], out);
      }
      continue;
    }

    if (node.children) walk(node.children, out);
    if (node.items) for (const item of node.items) walk(item.children, out);
  }
}

// ─── Decode ───────────────────────────────────────────────────────────────────

function decodeDataUri(uri) {
  if (!uri.startsWith('data:')) throw new Error('Unsupported image src (only data: URIs are supported)');
  const comma = uri.indexOf(',');
  if (comma === -1) throw new Error('Invalid data URI: missing comma');
  const header = uri.slice(5, comma); // strip leading "data:"
  const payload = uri.slice(comma + 1);
  if (header.endsWith(';base64')) return Buffer.from(payload, 'base64');
  return Buffer.from(decodeURIComponent(payload), 'utf8');
}

async function fetchHttpImage(url, options) {
  const timeoutMs = options.imageTimeoutMs ?? DEFAULT_IMAGE_TIMEOUT_MS;
  // No default cap — a caller who wants one for their own resource-safety
  // reasons (e.g. a public-facing service accepting arbitrary user HTML) can
  // still set options.maxImageBytes explicitly; nothing here imposes one
  // unasked. Infinity makes both size checks below no-ops without needing
  // their own separate "no limit configured" branch.
  const maxBytes = options.maxImageBytes ?? Infinity;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, { signal: controller.signal });
    if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);
    // A well-behaved server's own declared Content-Length lets an oversized
    // download get rejected before it's even read, not just after.
    const declaredLength = Number(response.headers.get('content-length'));
    if (Number.isFinite(declaredLength) && declaredLength > maxBytes) {
      throw new Error(`image exceeds maxImageBytes (${maxBytes})`);
    }
    const buffer = Buffer.from(await response.arrayBuffer());
    if (buffer.length > maxBytes) throw new Error(`image exceeds maxImageBytes (${maxBytes})`);
    return buffer;
  } catch (err) {
    if (err.name === 'AbortError') throw new Error(`timed out after ${timeoutMs}ms`);
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function makeId(src) {
  return createHash('sha1').update(src).digest('hex').slice(0, 16);
}

function guessMimeFromSrc(src) {
  const ext = src.split('?')[0]?.split('.').pop()?.toLowerCase() ?? '';
  const map = {
    jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png',
    gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml',
    bmp: 'image/bmp', ico: 'image/x-icon', avif: 'image/avif',
  };
  return map[ext] ?? 'application/octet-stream';
}
