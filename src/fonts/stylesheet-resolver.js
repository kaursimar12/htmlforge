/**
 * Provider-agnostic stylesheet discovery — the network/CSS-parsing half of
 * font resolution (see resolver.js's own doc comment for why this is kept
 * PDF-only and separate from the pure name-decision logic). Deliberately
 * contains NO special-casing for any specific provider (Google Fonts, Adobe
 * Fonts, a company CDN, ...): it only ever reads `<link rel="stylesheet">`
 * hrefs and `<style>`/`@import`-declared CSS text the HTML itself already
 * pointed at, via the exact same `@font-face`/`@import` grammar any of them
 * expose (see font-face-css.js). If a provider serves fonts through
 * standard CSS, this works automatically; nothing here ever needs to know
 * which provider is being talked to.
 *
 * Scoped to families the document actually asks for (collect-font-usages.js)
 * — a linked stylesheet declaring 20 weights only triggers font-FILE
 * downloads for the ones the document's own text actually uses; the
 * stylesheet's own (much smaller) CSS text is still fetched regardless,
 * since there's no way to know what it contains without reading it.
 *
 * Never throws — every fetch/parse/decode failure is caught per-entry and
 * turned into a warning, exactly like assets/resolver.js's own image
 * fetching; one unreachable stylesheet or corrupt font file never fails the
 * whole export, and the resolver's own Arial/Helvetica fallback (see
 * resolveFontFamilyStack) still applies for anything left unresolved.
 */
import { readFile } from 'node:fs/promises';
import { classifyImageSrc, resolveImageSrc } from '../utils/image-src.js';
import { parseFontFaceRules, parseImports, unicodeRangeCovers } from './font-face-css.js';
import { collectFontFamiliesUsed, collectCodepointsUsed } from './collect-font-usages.js';
import { decodeFontFile, cachedFetch } from './resolver.js';

const DEFAULT_TIMEOUT_MS = 10_000;
const DEFAULT_MAX_BYTES = 20 * 1024 * 1024;
const MAX_IMPORT_DEPTH = 1; // one level of @import chasing — see fetchStylesheetText's own doc note

// Module-level — persists for the lifetime of this process, across every
// convert() call, unlike the per-call `cache` object discoverFontsFromDocument
// builds below (which only dedupes within ONE call's own import/rule
// traversal). A remote stylesheet/font URL's bytes don't change in any way
// that matters within one running process, so a server converting many
// documents that share fonts only ever fetches each one once, total — not
// once per conversion. Deliberately keyed by URL only for HTTP(S) sources;
// a local file path is never cached here (see fetchText/fetchBytes below) —
// re-reading local disk is cheap, and caching it risks serving stale bytes
// if the same process converts multiple documents while a local font file
// changes on disk in between. A REJECTED fetch is evicted immediately (see
// cachedFetch) rather than remembered forever, so a transient network
// failure gets a real chance to succeed on a later, unrelated call.
const stylesheetTextCache = new Map();
const fontBytesCache = new Map();
// Caches the DECODED result (post WOFF2-transform-reconstruction, if any),
// not just the raw bytes fontBytesCache holds — decoding is real CPU work
// (this project's own triplet-decoding/glyph-reconstruction for a
// transformed WOFF2), a pure function of the bytes, so it's exactly as safe
// to reuse across calls as the raw fetch itself.
const decodedFontCache = new Map();

/** Clears every module-level cache above — production code never needs
 * this (the whole point is persisting for the process's life), but tests
 * that reuse the same mock URL across cases with different expected
 * responses need a way back to a clean slate between them. */
export function clearStylesheetCaches() {
  stylesheetTextCache.clear();
  fontBytesCache.clear();
  decodedFontCache.clear();
}

async function fetchText(url, options) {
  if (classifyImageSrc(url) !== 'http') return readFile(url, 'utf8');
  return cachedFetch(stylesheetTextCache, url, async () => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), options.fontTimeoutMs ?? DEFAULT_TIMEOUT_MS);
    try {
      const response = await fetch(url, { signal: controller.signal });
      if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);
      return await response.text();
    } finally {
      clearTimeout(timer);
    }
  });
}

async function fetchBytes(url, options) {
  if (classifyImageSrc(url) !== 'http') return readFile(url);
  const maxBytes = options.maxFontBytes ?? DEFAULT_MAX_BYTES;
  return cachedFetch(fontBytesCache, url, async () => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), options.fontTimeoutMs ?? DEFAULT_TIMEOUT_MS);
    try {
      const response = await fetch(url, { signal: controller.signal });
      if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);
      const declaredLength = Number(response.headers.get('content-length'));
      if (Number.isFinite(declaredLength) && declaredLength > maxBytes) throw new Error(`font exceeds maxFontBytes (${maxBytes})`);
      const buffer = Buffer.from(await response.arrayBuffer());
      if (buffer.length > maxBytes) throw new Error(`font exceeds maxFontBytes (${maxBytes})`);
      return buffer;
    } finally {
      clearTimeout(timer);
    }
  });
}

/** Fetches one stylesheet's CSS text plus, recursively, any `@import`ed
 * stylesheet's text (each import's own relative URLs — both nested
 * `@import`s and `@font-face src`s — resolved against THAT stylesheet's own
 * location, not the original page's) — capped at MAX_IMPORT_DEPTH to avoid
 * chasing an unbounded/circular import chain; deeper imports are simply not
 * followed rather than erroring, the same "bounded, not exhaustive" choice
 * this project already makes for e.g. descendant CSS selectors. Returns a
 * list of `{cssText, baseUrl}` — one entry per stylesheet successfully
 * fetched (the root one plus any followed imports), `baseUrl` being that
 * particular file's own location for resolving ITS OWN relative URLs.
 *
 * `cache.stylesheets` (keyed by resolved URL, shared across this whole
 * discoverFontsFromDocument call) makes a URL reachable more than once — the
 * same href linked twice, or the same stylesheet pulled in via two different
 * `@import` chains — fetch and get walked for further imports only the first
 * time; every later encounter reuses that same in-flight/settled promise. */
async function fetchStylesheetChain(url, options, warn, cache, depth = 0) {
  const results = [];
  if (!cache.stylesheets.has(url)) {
    cache.stylesheets.set(url, (async () => {
      try {
        return { cssText: await fetchText(url, options), baseUrl: url };
      } catch (err) {
        warn?.('STYLESHEET_FETCH_FAILED', `Cannot fetch stylesheet "${url}": ${err instanceof Error ? err.message : String(err)}`);
        return null;
      }
    })());
  } else {
    return results; // already fetched (and its own imports already walked) elsewhere in this call
  }

  const entry = await cache.stylesheets.get(url);
  if (!entry) return results;
  results.push(entry);

  if (depth < MAX_IMPORT_DEPTH) {
    for (const importUrl of parseImports(entry.cssText)) {
      const resolved = resolveImageSrc(importUrl, url);
      results.push(...await fetchStylesheetChain(resolved, options, warn, cache, depth + 1));
    }
  }
  return results;
}

/** Picks whichever Unicode-subset `@font-face` rule (candidates all share
 * the same family|weight|italic, differing only in unicode-range/src) covers
 * the most of `documentCodepoints` — the document's own actual text, not an
 * assumed script. A single candidate (the common case: most @font-face
 * declarations aren't subset-split at all) always wins trivially, no scoring
 * needed. When there's no real signal to score against (no codepoints
 * collected — e.g. the font is only ever referenced by a style with no
 * actual text under it — or every candidate genuinely covers none of the
 * document's text), falls back to the same "prefer whatever covers Basic
 * Latin, else the first declared" default this used unconditionally before:
 * Latin is still the single most common script when nothing else is known,
 * and "first declared" keeps this deterministic rather than picking
 * whichever network request happens to finish first. */
function pickBestSubsetCandidate(candidates, documentCodepoints) {
  if (candidates.length === 1) return candidates[0];

  let best;
  let bestScore = 0;
  for (const candidate of candidates) {
    let score = 0;
    for (const cp of documentCodepoints) {
      if (unicodeRangeCovers(candidate.rule.unicodeRange, cp)) score++;
    }
    if (score > bestScore) { bestScore = score; best = candidate; }
  }
  return best ?? candidates.find((c) => unicodeRangeCovers(c.rule.unicodeRange, 0x41)) ?? candidates[0];
}

/**
 * Discovers and registers fonts the document declares via `<link
 * rel="stylesheet">` and `<style>`/`@import`/`@font-face` — fills (and
 * returns) `registry`. This is the first font-resolution step index.js runs
 * (see its own call order), so `registry` starts empty here; later steps
 * (builtin-substitutes.js, the opt-in Google Fonts fallback) each check
 * `hasExact`/`has` before registering so they never override what this step
 * already resolved, the same "explicit/real beats generic assumption"
 * precedence this project already applies to margins/padding/etc.
 */
export async function discoverFontsFromDocument(document, options, warn, registry) {
  const familiesUsed = collectFontFamiliesUsed(document);
  if (familiesUsed.size === 0) return registry;

  // Scoped to this one call: `stylesheets` tracks which URLs have already
  // been fetched-and-walked-for-imports so a repeated href or shared
  // @import target isn't processed twice within this one document's own
  // discovery pass (the actual network fetch underneath is separately
  // cached at module level regardless — see fetchText above);
  // `localFontFiles` dedupes two @font-face rules in THIS call sharing the
  // same local file path down to one read+decode, without persisting past
  // this call (see the module-level decodedFontCache's own doc comment on
  // why a local path is never cached across calls).
  const cache = { stylesheets: new Map(), localFontFiles: new Map() };

  const stylesheetSources = [];

  const inlineCss = document.metadata?.inlineStylesheetCss;
  if (inlineCss) stylesheetSources.push({ cssText: inlineCss, baseUrl: options.baseUrl ?? process.cwd() });

  for (const href of document.metadata?.stylesheetLinks ?? []) {
    const resolved = resolveImageSrc(href, options.baseUrl ?? process.cwd());
    stylesheetSources.push(...await fetchStylesheetChain(resolved, options, warn, cache));
  }

  // Google Fonts/Bunny Fonts/etc. split ONE family+weight across several
  // @font-face rules by Unicode subset (`latin`, `latin-ext`, `cyrillic`,
  // `greek`, `vietnamese`...), each with its own DIFFERENT glyph coverage —
  // fetching all of them concurrently and registering whichever happens to
  // finish last (this project's own live testing against real Google/Bunny
  // stylesheets surfaced exactly this) would silently embed whichever
  // subset "won the race." Grouping by family|weight|italic and keeping only
  // the rule that covers the MOST of the document's own actual text
  // (collectCodepointsUsed — every candidate is scored by how many of the
  // document's real codepoints its unicode-range covers, not just assumed
  // Basic Latin) makes resolution deterministic and correct regardless of
  // script — a Cyrillic-primary or CJK-primary document picks the subset
  // that actually covers ITS OWN text, not whichever happens to include 'A'.
  const documentCodepoints = collectCodepointsUsed(document);
  const rulesByKey = new Map();
  for (const { cssText, baseUrl } of stylesheetSources) {
    for (const rule of parseFontFaceRules(cssText)) {
      if (!familiesUsed.has(rule.family.toLowerCase())) continue;
      if (registry.hasExact(rule.family, rule.weight, rule.italic)) continue; // e.g. the same rule reachable via two <link>s or an @import
      const key = `${rule.family.toLowerCase()}|${rule.weight}|${rule.italic}`;
      if (!rulesByKey.has(key)) rulesByKey.set(key, []);
      rulesByKey.get(key).push({ rule, baseUrl });
    }
  }

  const jobs = [];
  for (const candidates of rulesByKey.values()) {
    jobs.push(pickBestSubsetCandidate(candidates, documentCodepoints));
  }

  await Promise.all(jobs.map(async ({ rule, baseUrl }) => {
    for (const src of rule.srcs) {
      // `local(...)` sources (a system-installed font, e.g.
      // `src: local('Arial')`) have no bytes this project can fetch — the
      // cross-platform "locate an OS-installed font file" gap already
      // discussed is exactly why this project never attempts local(); it's
      // simply skipped in favor of the next source in the same rule.
      if (!src.url) continue;
      const resolvedUrl = resolveImageSrc(src.url, baseUrl);
      // Decoding (post WOFF2-transform-reconstruction if applicable) is a
      // pure function of the fetched bytes, so for a REMOTE source it's
      // cached at module level too (decodedFontCache), same "safe to reuse
      // for this process's whole lifetime" reasoning as fontBytesCache. A
      // local path is only deduped within THIS call (cache.localFontFiles)
      // instead — the whole point of not caching a local file's BYTES
      // across calls (see fetchBytes) is avoiding stale content if it
      // changes on disk between calls; caching the decode result at module
      // level would silently reintroduce that same staleness one step later.
      const decodePromise = classifyImageSrc(resolvedUrl) === 'http'
        ? cachedFetch(decodedFontCache, resolvedUrl, () => fetchBytes(resolvedUrl, options).then(decodeFontFile))
        : cachedFetch(cache.localFontFiles, resolvedUrl, () => fetchBytes(resolvedUrl, options).then(decodeFontFile));
      try {
        const { sfnt, buffer } = await decodePromise;
        // A variable font's glyf/loca only ever hold ONE default instance —
        // registering it under a non-regular weight/style would silently
        // embed IDENTICAL glyph data as a fake "Bold"/"Italic" (this project
        // has no variable-font instantiation engine to actually resolve a
        // real 700 or italic instance from it). Only the plain-regular case
        // is trustworthy from a variable source; any other weight/italic
        // combo resolving to one is refused so it falls through to the next
        // name in the CSS stack (or Arial) instead of claiming a fake match.
        if (sfnt.isVariable && (rule.weight !== 400 || rule.italic)) {
          warn?.('VARIABLE_FONT_WEIGHT_SKIPPED', `"${rule.family}" (weight ${rule.weight}${rule.italic ? ' italic' : ''}) resolved to a variable font — this project can't instantiate a real distinct weight/style from it, so this combo is left unresolved rather than embedding a duplicate of another weight.`);
          return;
        }
        registry.register(rule.family, rule.weight, rule.italic, sfnt, buffer);
        return;
      } catch (err) {
        // Try the next `src` entry in this same @font-face rule before
        // giving up on it entirely — a rule commonly lists a `.woff2` first
        // and a `.ttf` fallback, exactly for this situation.
        warn?.('FONT_FETCH_FAILED', `Cannot use font source "${src.url}" for "${rule.family}": ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  }));

  return registry;
}
