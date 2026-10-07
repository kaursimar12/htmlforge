/**
 * A small, purpose-built parser for the two @-rules a font resolver cares
 * about — `@font-face` (declares a family/weight/style + source URLs) and
 * `@import` (pulls in another stylesheet, possibly itself declaring more
 * `@font-face` rules) — out of raw CSS text (either a fetched stylesheet's
 * full body, or a `<style>` block's own text). Reuses
 * html-parser/inline-style.js's `parseInlineStyleAttr` for the declaration
 * block itself (`font-family: Poppins; font-weight: 700; src: url(...)...`)
 * since that's the exact same `prop: value; prop: value` grammar an inline
 * `style="..."` attribute already uses — no HTML-specific coupling in it,
 * so it applies unchanged to a CSS rule's own declaration body.
 */
import { parseInlineStyleAttr } from '../html-parser/inline-style.js';

const FONT_FACE_RE = /@font-face\s*\{([^{}]*)\}/g;
const IMPORT_RE = /@import\s+(?:url\(\s*["']?([^"')]+)["']?\s*\)|["']([^"']+)["'])[^;]*;?/g;
const SRC_ENTRY_RE = /url\(\s*["']?([^"')]+)["']?\s*\)(?:\s*format\(\s*["']?([^"')]+)["']?\s*\))?/g;

function extractFirstFontFamily(value) {
  if (!value) return undefined;
  const first = value.split(',')[0].trim().replace(/^['"]|['"]$/g, '');
  return first || undefined;
}

// `font-weight` in an @font-face rule may be a single value ("700") or,
// less commonly, a variable-font weight RANGE ("100 900") — the first
// number is used as this project's own single-weight model throughout
// doesn't represent a continuous variable-font range.
function parseWeight(value) {
  if (!value) return 400;
  if (value === 'bold') return 700;
  if (value === 'normal' || value === 'regular') return 400;
  const n = Number(value.trim().split(/\s+/)[0]);
  return Number.isFinite(n) ? n : 400;
}

// `U+XXXX` (single codepoint), `U+XXXX-YYYY` (range), and `U+XX??` (hex
// wildcard — each `?` matches any hex digit, i.e. `U+4E??` == U+4E00-U+4EFF)
// are CSS's own three unicode-range forms (used near-universally by Google
// Fonts/Bunny Fonts/etc. to split one family into a `latin`/`latin-ext`/
// `cyrillic`/`greek`/`vietnamese`/... rule per Unicode block, all sharing
// the SAME family+weight — see stylesheet-resolver.js's own doc comment on
// why picking the wrong one silently embeds a font with no Latin glyphs).
function parseUnicodeRange(value) {
  if (!value) return undefined;
  const ranges = [];
  for (const token of value.split(',')) {
    const trimmed = token.trim();
    const wildcard = /^U\+([0-9a-fA-F?]+)$/.exec(trimmed);
    if (wildcard && wildcard[1].includes('?')) {
      const hex = wildcard[1];
      const start = parseInt(hex.replace(/\?/g, '0'), 16);
      const end = parseInt(hex.replace(/\?/g, 'f'), 16);
      ranges.push([start, end]);
      continue;
    }
    const range = /^U\+([0-9a-fA-F]+)-([0-9a-fA-F]+)$/.exec(trimmed);
    if (range) { ranges.push([parseInt(range[1], 16), parseInt(range[2], 16)]); continue; }
    const single = /^U\+([0-9a-fA-F]+)$/.exec(trimmed);
    if (single) { const cp = parseInt(single[1], 16); ranges.push([cp, cp]); }
  }
  return ranges.length ? ranges : undefined;
}

/** Whether `unicodeRange` (as parsed above) covers `codePoint` — undefined
 * (no unicode-range declared at all) means "covers everything," matching
 * CSS's own semantics: no restriction means the font applies regardless of
 * character. */
export function unicodeRangeCovers(unicodeRange, codePoint) {
  if (!unicodeRange) return true;
  return unicodeRange.some(([start, end]) => codePoint >= start && codePoint <= end);
}

/** Every `@font-face` rule found in `cssText`, each with its own list of
 * `src: url(...) format(...)` entries (in declared order — a resolver tries
 * them in that order, the same "first workable source wins" a browser
 * itself applies to a multi-source @font-face declaration). */
export function parseFontFaceRules(cssText) {
  const rules = [];
  FONT_FACE_RE.lastIndex = 0;
  let m;
  while ((m = FONT_FACE_RE.exec(cssText))) {
    const decls = parseInlineStyleAttr(m[1]);
    const family = extractFirstFontFamily(decls['font-family']);
    if (!family) continue;
    const italic = decls['font-style'] === 'italic' || decls['font-style'] === 'oblique';
    const weight = parseWeight(decls['font-weight']);

    const srcs = [];
    SRC_ENTRY_RE.lastIndex = 0;
    let sm;
    while ((sm = SRC_ENTRY_RE.exec(decls['src'] ?? ''))) srcs.push({ url: sm[1], format: sm[2] });
    const unicodeRange = parseUnicodeRange(decls['unicode-range']);
    if (srcs.length) rules.push({ family, weight, italic, srcs, unicodeRange });
  }
  return rules;
}

/** Every `@import url(...)`/`@import "..."` target URL found in `cssText`. */
export function parseImports(cssText) {
  const urls = [];
  IMPORT_RE.lastIndex = 0;
  let m;
  while ((m = IMPORT_RE.exec(cssText))) urls.push(m[1] ?? m[2]);
  return urls;
}
