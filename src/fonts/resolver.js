/**
 * Font-family resolution — decides, for a CSS `font-family` fallback stack
 * plus a requested weight/italic, whether a caller-registered embedded font
 * (real bytes, parsed via sfnt-parser.js) exists for any name in the stack,
 * falling back to Arial/Helvetica when nothing in the stack has real glyph
 * data. Never throws — an unresolvable font is this module's normal "nothing
 * matched" case, not an error, matching this project's existing image-asset
 * "recoverable, warn, continue" discipline (see assets/resolver.js).
 *
 * Only PDF consumes the embedded-bytes half (`kind: 'embedded'`) today —
 * DOCX/PPTX just write a plain font-family name string and trust the
 * reading machine's own installed fonts (see their own w:rFonts/a:latin
 * write sites via resolveFontFamilyName below), so this module has two
 * genuinely separate resolution functions, not one shared implementation:
 * DOCX/PPTX can pass an arbitrary name through unchanged (the reading
 * machine supplies the glyphs), but PDF can only ever draw a font it has
 * actual bytes for, so an unregistered name is exactly as unusable as no
 * name at all.
 */
import { parseSfnt, isSfnt } from './sfnt-parser.js';
import { decodeWoff, isWoff } from './woff-decoder.js';
import { decodeWoff2, isWoff2 } from './woff2-decoder.js';

/**
 * Turns whatever font-file bytes were read/downloaded (`.ttf`/`.otf`/
 * `.woff`/`.woff2` — sniffed by magic bytes, not by file extension, the
 * same "trust the actual bytes over a possibly-wrong extension" precedent
 * `assets/resolver.js`'s own `detectFormat` follows for images) into a
 * plain sfnt buffer plus its parsed metadata. Shared by the local-file path
 * here and the future network-discovery path (stylesheet-declared fonts),
 * since both ultimately obtain the same four possible container formats.
 */
export function decodeFontFile(buffer) {
  const sfntBuffer = isWoff(buffer) ? decodeWoff(buffer)
    : isWoff2(buffer) ? decodeWoff2(buffer)
    : buffer;
  if (!isSfnt(sfntBuffer)) throw new Error('not a recognized TrueType/OpenType/WOFF/WOFF2 font file');
  return { sfnt: parseSfnt(sfntBuffer), buffer: sfntBuffer };
}

/**
 * Memoizes `fetchFn()`'s result in `cache` under `key`, for as long as
 * `cache` itself lives — shared by stylesheet-resolver.js (module-level
 * `cache`, so a URL's bytes are fetched at most once per process, not once
 * per convert() call) and google-fonts-fallback.js (identical reasoning, its
 * own separate module-level caches). A REJECTED fetch is evicted immediately
 * rather than remembered forever, so a transient network failure gets a real
 * chance to succeed on a later, unrelated call instead of being permanently
 * remembered as broken for the rest of the process's lifetime.
 */
export function cachedFetch(cache, key, fetchFn) {
  if (!cache.has(key)) {
    const promise = fetchFn().catch((err) => {
      cache.delete(key);
      throw err;
    });
    cache.set(key, promise);
  }
  return cache.get(key);
}

const FALLBACK_NAME = 'Arial';

function splitFontFamilyStack(value) {
  if (!value) return [];
  return value.split(',').map((entry) => entry.trim().replace(/^['"]|['"]$/g, '')).filter(Boolean);
}

function normalizeWeightKey(weight, italic) {
  return `${weight}${italic ? 'i' : ''}`;
}

/**
 * Registry of embedded fonts resolved from a document's own declared
 * stylesheets/@font-face (stylesheet-resolver.js), the 11 bundled fonts
 * (builtin-substitutes.js), and the opt-in Google Fonts fallback
 * (google-fonts-fallback.js) — one per convert() call, never persisted
 * across calls (matches assets/resolver.js's own per-conversion `seen` Map
 * dedup, not a cross-process cache).
 */
export class EmbeddedFontRegistry {
  #familiesLowercase = new Map(); // lowercased family -> Map<weightKey, {sfnt, buffer, family, weight, italic}>

  register(family, weight, italic, sfnt, buffer) {
    const key = family.toLowerCase();
    if (!this.#familiesLowercase.has(key)) this.#familiesLowercase.set(key, new Map());
    this.#familiesLowercase.get(key).set(normalizeWeightKey(weight, italic), { sfnt, buffer, family, weight, italic });
  }

  has(family) {
    return this.#familiesLowercase.has(family.toLowerCase());
  }

  /** True only for an EXACT weight/italic combo already registered — unlike
   * `match()` (which always finds the nearest available weight once any
   * variant of `family` exists), this is what a caller needs to ask "has
   * THIS specific variant already been resolved" (e.g. builtin-substitutes.js
   * deciding whether a document-discovered @font-face already covers a
   * weight before loading its own bundled substitute for the same name). */
  hasExact(family, weight, italic) {
    return Boolean(this.#familiesLowercase.get(family.toLowerCase())?.has(normalizeWeightKey(weight, italic)));
  }

  /** Best available entry for `family` at the requested weight/italic — exact
   * match preferred, else same-italic nearest weight, else nearest weight
   * regardless of italic (approximation, same "closest available" spirit as
   * this file's other fallbacks). Returns undefined if `family` was never
   * registered at all. */
  match(family, weight, italic) {
    const entries = this.#familiesLowercase.get(family.toLowerCase());
    if (!entries || entries.size === 0) return undefined;

    const exact = entries.get(normalizeWeightKey(weight, italic));
    if (exact) return exact;

    let best;
    let bestScore = Infinity;
    for (const entry of entries.values()) {
      const italicPenalty = entry.italic === italic ? 0 : 1000;
      const score = italicPenalty + Math.abs(entry.weight - weight);
      if (score < bestScore) { bestScore = score; best = entry; }
    }
    return best;
  }

  get size() { return this.#familiesLowercase.size; }
}

// DOCX/PPTX's own generic-keyword map, using real installable Windows/Mac
// font names instead of PDF Standard-14 names ("Helvetica"/"Times-Roman"/
// "Courier" aren't real fonts anyone has installed — Word/PowerPoint would
// just fail to find them and silently substitute something anyway).
// Deliberately NOT the same resolution function as resolveFontFamilyStack
// below: DOCX/PPTX never embed bytes, so an arbitrary custom name (e.g.
// "Poppins") should pass through unchanged exactly as it does today — only a
// bare generic keyword (which was never a real font name to begin with)
// gets mapped; nothing else ever falls back to Arial the way an unresolvable
// PDF run would.
const GENERIC_TO_WEB_SAFE_NAME = {
  serif: 'Times New Roman',
  'sans-serif': 'Arial',
  monospace: 'Courier New',
  cursive: 'Arial',
  fantasy: 'Arial',
};

/**
 * DOCX/PPTX's font-name resolution — maps a bare CSS generic keyword to a
 * real installable name; any other value (a web-safe name or an arbitrary
 * custom family like "Poppins") is returned unchanged. No network, no
 * embedding, never fails — matches DOCX/PPTX's own "just write a name,
 * trust the reading machine" behavior, only fixing the one case (a literal
 * generic keyword) that was never a valid font name to write as-is.
 */
export function resolveFontFamilyName(fontFamily) {
  if (!fontFamily) return fontFamily;
  const lower = fontFamily.trim().toLowerCase();
  return GENERIC_TO_WEB_SAFE_NAME[lower] ?? fontFamily;
}

/**
 * Resolves one CSS `font-family` fallback stack (e.g. `"Poppins", Arial,
 * sans-serif`) to a definitive choice — searches every entry, left to right,
 * for a REAL registered font (a page-discovered `@font-face`, a bundled
 * font — see fonts/builtin-substitutes.js, which registers e.g. "Arial"
 * under real Arimo glyph data — or the opt-in Google Fonts fallback) before
 * giving up. There's no name-only "assume it's
 * installed and guess the closest Standard-14 family" bucket for a generic
 * keyword or a web-safe name that never resolved to real glyph data: PDF
 * output can only ever draw a font it actually has bytes for, so a name with
 * no registry match is exactly as unresolved as one further down the stack —
 * checking every entry for a real match (rather than stopping early at the
 * first "plausible" name) means a real embedded font later in the stack
 * still wins over an earlier name we could only ever guess at anyway.
 * Returns one of:
 *   { kind: 'embedded', name, entry }    — entry is the registry match
 *   { kind: 'fallback', name: 'Arial' }  — nothing in the stack resolved; never fails
 */
export function resolveFontFamilyStack(fontFamilyCss, { weight = 400, italic = false } = {}, registry) {
  for (const rawName of splitFontFamilyStack(fontFamilyCss)) {
    const entry = registry?.match(rawName, weight, italic);
    if (entry) return { kind: 'embedded', name: rawName, entry };
  }
  return { kind: 'fallback', name: FALLBACK_NAME };
}
