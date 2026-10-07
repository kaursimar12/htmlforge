/**
 * Opt-in, deliberately provider-specific exception to this project's
 * otherwise provider-agnostic font resolution (see stylesheet-resolver.js's
 * own doc comment on why that file only ever follows a URL the HTML itself
 * points to). This module instead GUESSES a URL — Google Fonts' public CSS2
 * API, by family name alone — for whatever families the document asks for
 * but never declared a source for at all (no `@font-face` discovered, not
 * one of the 11 bundled fonts). Gated behind
 * `options.googleFontsFallback` (default false) so it never makes a silent
 * network call to one specific third party unless the caller asked for it —
 * unlike stylesheet-resolver.js's own discovery (unconditional, since it
 * only ever follows a URL the document itself already named), this one
 * fetches somewhere the document never pointed, so it stays opt-in.
 *
 * Reuses the exact same CSS-parsing/font-decoding machinery
 * stylesheet-resolver.js already built for document-declared @font-face —
 * the Google Fonts CSS2 response IS just another stylesheet to this code;
 * nothing here is Google-specific below the URL-building step itself.
 */
import { parseFontFaceRules, unicodeRangeCovers } from './font-face-css.js';
import { decodeFontFile, cachedFetch } from './resolver.js';
import { collectFontFamiliesUsed } from './collect-font-usages.js';

const DEFAULT_TIMEOUT_MS = 10_000;
const DEFAULT_MAX_BYTES = 20 * 1024 * 1024;

// Module-level — persists for this process's whole lifetime, across every
// convert() call, same reasoning as stylesheet-resolver.js's own identical
// caches: Google's response for a given family/URL doesn't change in any way
// that matters here, so a server resolving the same undeclared font across
// many conversions only ever hits Google once, total. `decodedFontCache`
// caches the DECODED result (not just raw bytes) since decoding is real CPU
// work (WOFF2-transform reconstruction, when applicable) and a pure function
// of the bytes — safe to reuse for as long as the bytes are. A REJECTED
// fetch is evicted immediately (see resolver.js's cachedFetch) so a
// transient failure gets a real retry chance on a later, unrelated call.
const cssTextCache = new Map();
const decodedFontCache = new Map();

/** Clears both module-level caches above — see stylesheet-resolver.js's
 * identical `clearStylesheetCaches` for why this exists (test isolation
 * when reusing a mock URL across cases; production code never calls it). */
export function clearGoogleFontsCache() {
  cssTextCache.clear();
  decodedFontCache.clear();
}

function buildCss2Url(family) {
  return `https://fonts.googleapis.com/css2?family=${encodeURIComponent(family)}:wght@400;700&display=swap`;
}

// Google's CSS2 API is case-sensitive (family=Lato resolves; family=lato/
// family=LATO both 400) and its real catalog names are conventionally
// Title Case — but not always word-by-word (e.g. "IBM Plex Sans" keeps
// "IBM" all-caps), so this is only ever tried as a SECOND attempt, after
// the CSS-declared casing verbatim already failed: a correctly-cased
// unusual name still succeeds on the first try and never reaches this,
// while a simple miscasing (`LATO`, `open sans`) gets a real second chance
// instead of silently failing outright.
function toTitleCase(family) {
  return family.toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());
}

async function fetchCssText(url, options) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), options.fontTimeoutMs ?? DEFAULT_TIMEOUT_MS);
  try {
    // Deliberately NOT spoofing a modern browser User-Agent (unlike a real
    // browser would send) — the CSS2 endpoint's response is entirely
    // different depending on this. A modern UA gets today's actual browser
    // delivery: one variable-font file per family, shared across every
    // requested weight/style, split across several Unicode-subset rules —
    // exactly the "same file registered under two different weight keys"
    // problem this project has no instantiation engine to safely resolve
    // (see the `isVariable` check below, kept as a safety net for other
    // providers). An unrecognized/default UA instead gets Google's legacy
    // fallback: one plain, genuinely-distinct static `.ttf` per weight/style,
    // no subsetting, no variable axes at all — simpler AND more correct for
    // this project's purposes, so this is the deliberate choice, not an
    // oversight.
    const response = await fetch(url, { signal: controller.signal });
    if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);
    return await response.text();
  } finally {
    clearTimeout(timer);
  }
}

async function fetchFontBytes(url, options) {
  const maxBytes = options.maxFontBytes ?? DEFAULT_MAX_BYTES;
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
}

/**
 * Tries Google Fonts' CSS2 API, by family name alone, for every family the
 * document uses that isn't already registered by any earlier source
 * (discovered `@font-face`, or one of the 11 bundled fonts — this only ever
 * runs AFTER those, in index.js). Fills (and
 * returns) `registry`, same contract as this module's siblings. Never
 * throws: a family Google doesn't have (its CSS2 endpoint 404s for an
 * unknown name), a network failure, or a corrupt font file is warned and
 * skipped, leaving that family to fall through to the next name in its own
 * CSS fallback stack (or Arial) exactly as if this step had never run.
 */
export async function fetchGoogleFontsFallback(document, options, warn, registry) {
  // .values() specifically — collectFontFamiliesUsed's map values are the
  // original-cased names (its keys are lowercased for other callers); the
  // original casing matters here since Google's CSS2 API is case-sensitive.
  const families = [...collectFontFamiliesUsed(document).values()].filter((family) => !registry.has(family));
  if (families.length === 0) return registry;

  await Promise.all(families.map(async (family) => {
    // Try the CSS-declared casing first; only retry with Title Case if that
    // fails, and only when it's actually a different string (skips a
    // wasted identical second request for a name already in Title Case).
    const candidateNames = [family, toTitleCase(family)].filter((name, i, arr) => arr.indexOf(name) === i);

    let cssText;
    let lastErr;
    for (const candidateName of candidateNames) {
      const url = buildCss2Url(candidateName);
      try {
        cssText = await cachedFetch(cssTextCache, url, () => fetchCssText(url, options));
        lastErr = undefined;
        break;
      } catch (err) {
        lastErr = err;
      }
    }
    if (lastErr) {
      warn?.('GOOGLE_FONT_LOOKUP_FAILED', `Cannot find "${family}" via Google Fonts fallback: ${lastErr instanceof Error ? lastErr.message : String(lastErr)}`);
      return;
    }

    // Same "a provider splits one family+weight across several Unicode
    // subset rules" situation stylesheet-resolver.js already handles for
    // document-declared @font-face — identical fix applies here since this
    // is the exact same CSS2 API.
    const rulesByKey = new Map();
    for (const rule of parseFontFaceRules(cssText)) {
      const key = `${rule.weight}|${rule.italic}`;
      if (!rulesByKey.has(key)) rulesByKey.set(key, []);
      rulesByKey.get(key).push(rule);
    }

    await Promise.all([...rulesByKey.values()].map(async (candidates) => {
      const rule = candidates.find((r) => unicodeRangeCovers(r.unicodeRange, 0x41)) ?? candidates[0];
      for (const src of rule.srcs) {
        if (!src.url) continue;
        try {
          const { sfnt, buffer } = await cachedFetch(decodedFontCache, src.url, () => fetchFontBytes(src.url, options).then(decodeFontFile));
          // See stylesheet-resolver.js's identical check — a variable font's
          // glyf/loca is one default instance; only the plain-regular case
          // is trustworthy without a real instantiation engine.
          if (sfnt.isVariable && (rule.weight !== 400 || rule.italic)) {
            warn?.('VARIABLE_FONT_WEIGHT_SKIPPED', `"${family}" (weight ${rule.weight}${rule.italic ? ' italic' : ''}) resolved to a variable font via Google Fonts — this project can't instantiate a real distinct weight/style from it, so this combo is left unresolved rather than embedding a duplicate of another weight.`);
            return;
          }
          registry.register(family, rule.weight, rule.italic, sfnt, buffer);
          return;
        } catch (err) {
          warn?.('GOOGLE_FONT_FETCH_FAILED', `Cannot use Google Fonts source for "${family}": ${err instanceof Error ? err.message : String(err)}`);
        }
      }
    }));
  }));

  return registry;
}
