/**
 * HTML/XML text-escaping helpers plus a text-safe image-src fallback — shared
 * wherever an exporter needs to work with raw HTML/XML strings directly.
 */

import { resolveImageSrc } from './image-src.js';

/** Escape a string for safe inclusion in HTML text content. */
export function escapeHtml(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Escape a string for safe inclusion inside a double-quoted HTML attribute. */
export function escapeAttribute(value) {
  return escapeHtml(value).replace(/"/g, '&quot;');
}

/**
 * A safe, human-readable form of an image src for text-based output formats
 * (Markdown/TXT/CSV/JSON's content tree) — a data: URI can be megabytes of
 * base64, which is never useful to print verbatim inline. Real URLs and file
 * paths pass through unchanged, except a relative one is resolved against
 * `baseUrl` via the exact same pure resolution the asset resolver itself
 * uses before ever fetching/reading bytes (utils/image-src.js's own
 * resolveImageSrc) — one shared implementation, so a text exporter's printed
 * reference and a binary exporter's actual fetch always agree on what a
 * given relative src resolves to. Returns undefined for a data: URI (or no
 * src at all), so callers fall back to alt text / a generic placeholder the
 * same way they already do when no src is present. Text exporters never
 * call the asset resolver itself — this is resolution only, never a fetch or
 * a file read.
 */
export function displayImageSrc(src, baseUrl) {
  if (!src || src.startsWith('data:')) return undefined;
  return resolveImageSrc(src, baseUrl);
}
