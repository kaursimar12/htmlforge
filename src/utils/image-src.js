/**
 * Pure image-src classification/resolution — string/URL/path manipulation
 * only, no filesystem or network access. This is the single source of truth
 * for "what does this src actually point at," shared by the asset resolver
 * (which uses it before fetching/reading real bytes — see
 * assets/resolver.js) and the text exporters (Markdown/TXT/JSON), which need
 * the exact same relative→absolute resolution to print a readable reference
 * but must never perform any actual image loading themselves.
 */

import path from 'node:path';

/** `'data' | 'http' | 'local'` — never inspects bytes, only the src string itself. */
export function classifyImageSrc(src) {
  if (src.startsWith('data:')) return 'data';
  if (/^https?:\/\//i.test(src)) return 'http';
  return 'local';
}

/**
 * A relative `src` ('local' — no data:/http(s):// prefix) resolves against
 * `baseUrl`: as a URL base when baseUrl itself is an http(s) URL (matches
 * the PRD's `baseUrl: "https://myapp.com"` example — a relative image path
 * on a remote page), otherwise as a filesystem directory. `data:`/
 * `http(s)://` srcs are already absolute and pass through unchanged
 * regardless of `baseUrl`; a `local` src with no `baseUrl` at all also
 * passes through unchanged — there's no base to resolve it against, so it's
 * left exactly as authored rather than guessing one. (The asset resolver,
 * which needs a concrete root to actually open a real file, supplies
 * `process.cwd()` itself as the fallback `baseUrl` when the caller gave none
 * — see assets/resolver.js — rather than this pure function assuming that
 * policy on every caller's behalf.)
 */
export function resolveImageSrc(src, baseUrl) {
  if (classifyImageSrc(src) !== 'local' || !baseUrl) return src;
  if (/^https?:\/\//i.test(baseUrl)) return new URL(src, baseUrl).toString();
  return path.resolve(baseUrl, src);
}
