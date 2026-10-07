/**
 * Font metrics — dependency-free text width/wrap estimation for DrawingML
 * shapes, shared by both the PPTX and DOCX exporters.
 *
 * Neither PowerPoint's DrawingML paragraphs nor a Word paragraph promoted to
 * a real `wps:wsp` shape (see pptx/shapes.js's renderStyledBoxShape/
 * renderCardShape and docx/document.js's renderCircularBadgeShape/
 * renderCardShapeDrawing) has any way to ask "how wide will this text
 * actually render" — so sizing a badge/card/circle correctly means guessing,
 * with no DOM, canvas, or font file available (zero runtime deps).
 *
 * The named tables below are per-mille (‰, i.e. /1000 = em fraction) advance
 * widths for the ASCII printable range (0x20–0x7E). Every one of them is
 * generated — not hand-typed or scaled — from the real, OFL-licensed,
 * metric-compatible substitute fonts bundled at src/fonts/builtin/ (Arimo,
 * Tinos, Cousine, Gelasio, Carlito — see fonts/builtin-substitutes.js for why
 * those specific fonts stand in for Arial/Times New Roman/Courier New/
 * Georgia/Calibri) via scripts/generate-font-metrics.mjs, which reads each
 * font's own `hmtx` table through sfnt-parser.js and prints the resulting
 * object literals. Re-run that script and paste its output here if the
 * bundled fonts ever change; this module itself stays pure sync/zero-I/O —
 * it never reads a font file, only these checked-in numbers. (For Arial/
 * Helvetica/Times New Roman this reproduces the well-known Adobe Core-14
 * values exactly, confirming both the generation pipeline and the fonts'
 * own metric-compatibility claim; Georgia and Calibri previously had no real
 * table at all and were a scaled guess off the nearest compatible font —
 * this replaces that guess with real measured data.)
 */

// Generated from Arimo-Regular.ttf (Arial/Helvetica's bundled substitute) via
// scripts/generate-font-metrics.mjs — do not hand-edit. Matches the
// published Adobe Core-14 "Helvetica" widths exactly.
// prettier-ignore
const HELVETICA_WIDTHS = {
  ' ': 278, '!': 278, '"': 355, '#': 556, '$': 556, '%': 889, '&': 667, "'": 191,
  '(': 333, ')': 333, '*': 389, '+': 584, ',': 278, '-': 333, '.': 278, '/': 278,
  '0': 556, '1': 556, '2': 556, '3': 556, '4': 556, '5': 556, '6': 556, '7': 556, '8': 556, '9': 556,
  ':': 278, ';': 278, '<': 584, '=': 584, '>': 584, '?': 556, '@': 1015,
  A: 667, B: 667, C: 722, D: 722, E: 667, F: 611, G: 778, H: 722, I: 278, J: 500,
  K: 667, L: 556, M: 833, N: 722, O: 778, P: 667, Q: 778, R: 722, S: 667, T: 611,
  U: 722, V: 667, W: 944, X: 667, Y: 667, Z: 611,
  '[': 278, '\\': 278, ']': 278, '^': 469, _: 556, '`': 333,
  a: 556, b: 556, c: 500, d: 556, e: 556, f: 278, g: 556, h: 556, i: 222, j: 222,
  k: 500, l: 222, m: 833, n: 556, o: 556, p: 556, q: 556, r: 333, s: 500, t: 278,
  u: 556, v: 500, w: 722, x: 500, y: 500, z: 500,
  '{': 334, '|': 260, '}': 334, '~': 584,
};

// Generated from Arimo-Bold.ttf via scripts/generate-font-metrics.mjs — do
// not hand-edit. A bold run actually draws as the real, wider "Helvetica-
// Bold" Standard-14 font (see pdf/fonts.js's pickBaseFont) — measuring it
// against the REGULAR table above under-estimates its true width, which is
// exactly what let a bold run's drawn glyphs run past where the next
// (differently-styled) segment was positioned to start, visually colliding
// with it — e.g. "**Web-safe** (Georgia)" losing its gap entirely. This
// table (and its TIMES_BOLD_WIDTHS/COURIER_BOLD_WIDTHS/GEORGIA_BOLD_WIDTHS/
// CALIBRI_BOLD_WIDTHS siblings below) exists so `estimateTextWidth`'s own
// `bold` parameter has real, measured bold-weight data to consult instead of
// silently reusing the regular table for every weight.
// prettier-ignore
const HELVETICA_BOLD_WIDTHS = {
  ' ': 278, '!': 333, '"': 474, '#': 556, '$': 556, '%': 889, '&': 722, "'": 238,
  '(': 333, ')': 333, '*': 389, '+': 584, ',': 278, '-': 333, '.': 278, '/': 278,
  '0': 556, '1': 556, '2': 556, '3': 556, '4': 556, '5': 556, '6': 556, '7': 556, '8': 556, '9': 556,
  ':': 333, ';': 333, '<': 584, '=': 584, '>': 584, '?': 611, '@': 975,
  A: 722, B: 722, C: 722, D: 722, E: 667, F: 611, G: 778, H: 722, I: 278, J: 556,
  K: 722, L: 611, M: 833, N: 722, O: 778, P: 667, Q: 778, R: 722, S: 667, T: 611,
  U: 722, V: 667, W: 944, X: 667, Y: 667, Z: 611,
  '[': 333, '\\': 278, ']': 333, '^': 584, _: 556, '`': 333,
  a: 556, b: 611, c: 556, d: 611, e: 556, f: 333, g: 611, h: 611, i: 278, j: 278,
  k: 556, l: 278, m: 889, n: 611, o: 611, p: 611, q: 611, r: 389, s: 556, t: 333,
  u: 611, v: 556, w: 778, x: 556, y: 556, z: 500,
  '{': 389, '|': 280, '}': 389, '~': 584,
};

// Generated from Tinos-Regular.ttf (Times New Roman's bundled substitute)
// via scripts/generate-font-metrics.mjs — do not hand-edit. Matches the
// published Adobe Core-14 "Times-Roman" widths exactly.
// prettier-ignore
const TIMES_WIDTHS = {
  ' ': 250, '!': 333, '"': 408, '#': 500, '$': 500, '%': 833, '&': 778, "'": 180,
  '(': 333, ')': 333, '*': 500, '+': 564, ',': 250, '-': 333, '.': 250, '/': 278,
  '0': 500, '1': 500, '2': 500, '3': 500, '4': 500, '5': 500, '6': 500, '7': 500, '8': 500, '9': 500,
  ':': 278, ';': 278, '<': 564, '=': 564, '>': 564, '?': 444, '@': 921,
  A: 722, B: 667, C: 667, D: 722, E: 611, F: 556, G: 722, H: 722, I: 333, J: 389,
  K: 722, L: 611, M: 889, N: 722, O: 722, P: 556, Q: 722, R: 667, S: 556, T: 611,
  U: 722, V: 722, W: 944, X: 722, Y: 722, Z: 611,
  '[': 333, '\\': 278, ']': 333, '^': 469, _: 500, '`': 333,
  a: 444, b: 500, c: 444, d: 500, e: 444, f: 333, g: 500, h: 500, i: 278, j: 278,
  k: 500, l: 278, m: 778, n: 500, o: 500, p: 500, q: 500, r: 333, s: 389, t: 278,
  u: 500, v: 500, w: 722, x: 500, y: 500, z: 444,
  '{': 480, '|': 200, '}': 480, '~': 541,
};

// Generated from Tinos-Bold.ttf via scripts/generate-font-metrics.mjs — do
// not hand-edit. Bold sibling of TIMES_WIDTHS — see HELVETICA_BOLD_WIDTHS'
// own comment for why a dedicated bold table (not the regular one) matters.
// prettier-ignore
const TIMES_BOLD_WIDTHS = {
  ' ': 250, '!': 333, '"': 555, '#': 500, '$': 500, '%': 1000, '&': 833, "'": 278,
  '(': 333, ')': 333, '*': 500, '+': 570, ',': 250, '-': 333, '.': 250, '/': 278,
  '0': 500, '1': 500, '2': 500, '3': 500, '4': 500, '5': 500, '6': 500, '7': 500, '8': 500, '9': 500,
  ':': 333, ';': 333, '<': 570, '=': 570, '>': 570, '?': 500, '@': 930,
  A: 722, B: 667, C: 722, D: 722, E: 667, F: 611, G: 778, H: 778, I: 389, J: 500,
  K: 778, L: 667, M: 944, N: 722, O: 778, P: 611, Q: 778, R: 722, S: 556, T: 667,
  U: 722, V: 722, W: 1000, X: 722, Y: 722, Z: 667,
  '[': 333, '\\': 278, ']': 333, '^': 581, _: 500, '`': 333,
  a: 500, b: 556, c: 444, d: 556, e: 444, f: 333, g: 500, h: 556, i: 278, j: 333,
  k: 556, l: 278, m: 833, n: 556, o: 500, p: 556, q: 556, r: 444, s: 389, t: 333,
  u: 556, v: 500, w: 722, x: 500, y: 500, z: 444,
  '{': 394, '|': 220, '}': 394, '~': 520,
};

// Generated from Gelasio-Regular.ttf (Georgia's bundled substitute) via
// scripts/generate-font-metrics.mjs — do not hand-edit. Real measured data;
// previously this was a scaled guess off TIMES_WIDTHS with no real font
// behind it at all.
// prettier-ignore
const GEORGIA_WIDTHS = {
  ' ': 241, '!': 331, '"': 412, '#': 643, '$': 610, '%': 817, '&': 710, "'": 215,
  '(': 375, ')': 375, '*': 472, '+': 643, ',': 270, '-': 374, '.': 270, '/': 469,
  '0': 614, '1': 430, '2': 559, '3': 552, '4': 565, '5': 528, '6': 566, '7': 502, '8': 596, '9': 566,
  ':': 313, ';': 313, '<': 643, '=': 643, '>': 643, '?': 479, '@': 929,
  A: 671, B: 654, C: 642, D: 749, E: 653, F: 599, G: 725, H: 815, I: 390, J: 518,
  K: 694, L: 604, M: 927, N: 767, O: 744, P: 610, Q: 744, R: 702, S: 561, T: 619,
  U: 756, V: 667, W: 976, X: 710, Y: 615, Z: 602,
  '[': 375, '\\': 469, ']': 375, '^': 643, _: 643, '`': 500,
  a: 504, b: 560, c: 454, d: 574, e: 483, f: 325, g: 509, h: 582, i: 293, j: 292,
  k: 536, l: 286, m: 881, n: 591, o: 539, p: 571, q: 560, r: 410, s: 432, t: 345,
  u: 575, v: 497, w: 737, x: 505, y: 492, z: 444,
  '{': 430, '|': 375, '}': 430, '~': 643,
};

// Generated from Gelasio-Bold.ttf via scripts/generate-font-metrics.mjs — do
// not hand-edit. Bold sibling of GEORGIA_WIDTHS.
// prettier-ignore
const GEORGIA_BOLD_WIDTHS = {
  ' ': 254, '!': 376, '"': 510, '#': 703, '$': 641, '%': 879, '&': 799, "'": 269,
  '(': 447, ')': 447, '*': 482, '+': 703, ',': 328, '-': 379, '.': 328, '/': 472,
  '0': 701, '1': 490, '2': 626, '3': 625, '4': 649, '5': 599, '6': 648, '7': 554, '8': 676, '9': 648,
  ':': 367, ';': 367, '<': 703, '=': 703, '>': 703, '?': 548, '@': 967,
  A: 758, B: 757, C: 715, D: 834, E: 721, F: 671, G: 807, H: 913, I: 446, J: 595,
  K: 817, L: 686, M: 1023, N: 839, O: 820, P: 701, Q: 820, R: 797, S: 649, T: 684,
  U: 833, V: 762, W: 1126, X: 809, Y: 732, Z: 689,
  '[': 447, '\\': 472, ']': 447, '^': 703, _: 703, '`': 500,
  a: 596, b: 646, c: 531, d: 663, e: 572, f: 393, g: 577, h: 680, i: 354, j: 346,
  k: 632, l: 344, m: 1016, n: 690, o: 636, p: 658, q: 648, r: 520, s: 513, t: 397,
  u: 677, v: 567, w: 863, x: 588, y: 562, z: 525,
  '{': 500, '|': 388, '}': 500, '~': 703,
};

// Generated from Carlito-Regular.ttf (Calibri's bundled substitute) via
// scripts/generate-font-metrics.mjs — do not hand-edit. Real measured data;
// previously this was a scaled guess off HELVETICA_WIDTHS with no real font
// behind it at all.
// prettier-ignore
const CALIBRI_WIDTHS = {
  ' ': 226, '!': 326, '"': 401, '#': 498, '$': 507, '%': 715, '&': 682, "'": 221,
  '(': 303, ')': 303, '*': 498, '+': 498, ',': 250, '-': 306, '.': 252, '/': 386,
  '0': 507, '1': 507, '2': 507, '3': 507, '4': 507, '5': 507, '6': 507, '7': 507, '8': 507, '9': 507,
  ':': 268, ';': 268, '<': 498, '=': 498, '>': 498, '?': 463, '@': 894,
  A: 579, B: 544, C: 533, D: 615, E: 488, F: 459, G: 631, H: 623, I: 252, J: 319,
  K: 520, L: 420, M: 855, N: 646, O: 662, P: 517, Q: 673, R: 543, S: 459, T: 487,
  U: 642, V: 567, W: 890, X: 519, Y: 487, Z: 468,
  '[': 307, '\\': 386, ']': 307, '^': 498, _: 498, '`': 291,
  a: 479, b: 525, c: 423, d: 525, e: 498, f: 305, g: 471, h: 525, i: 229, j: 239,
  k: 455, l: 229, m: 799, n: 525, o: 527, p: 525, q: 525, r: 349, s: 391, t: 335,
  u: 525, v: 452, w: 715, x: 433, y: 453, z: 395,
  '{': 314, '|': 460, '}': 314, '~': 498,
};

// Generated from Carlito-Bold.ttf via scripts/generate-font-metrics.mjs — do
// not hand-edit. Bold sibling of CALIBRI_WIDTHS.
// prettier-ignore
const CALIBRI_BOLD_WIDTHS = {
  ' ': 226, '!': 326, '"': 438, '#': 498, '$': 507, '%': 729, '&': 705, "'": 233,
  '(': 312, ')': 312, '*': 498, '+': 498, ',': 258, '-': 306, '.': 267, '/': 430,
  '0': 507, '1': 507, '2': 507, '3': 507, '4': 507, '5': 507, '6': 507, '7': 507, '8': 507, '9': 507,
  ':': 276, ';': 276, '<': 498, '=': 498, '>': 498, '?': 463, '@': 898,
  A: 606, B: 561, C: 529, D: 630, E: 488, F: 459, G: 637, H: 631, I: 267, J: 331,
  K: 547, L: 423, M: 874, N: 659, O: 676, P: 532, Q: 686, R: 563, S: 473, T: 495,
  U: 653, V: 591, W: 906, X: 551, Y: 520, Z: 478,
  '[': 325, '\\': 430, ']': 325, '^': 498, _: 498, '`': 300,
  a: 494, b: 537, c: 418, d: 537, e: 503, f: 316, g: 474, h: 537, i: 246, j: 255,
  k: 480, l: 246, m: 813, n: 537, o: 538, p: 537, q: 537, r: 355, s: 399, t: 347,
  u: 537, v: 473, w: 745, x: 459, y: 474, z: 397,
  '{': 344, '|': 475, '}': 344, '~': 498,
};

// Generated from Cousine-Regular.ttf (Courier New's bundled substitute) via
// scripts/generate-font-metrics.mjs, confirming what monospace already
// guarantees: every glyph advances exactly 600‰ of an em, printable range
// included, matching the published Adobe Core-14 Courier AFM too. Built as a
// flat table (rather than a special-cased "if courier, return 0.6" branch in
// charWidthEm) so it composes with the existing named-table lookup path
// unchanged. Cousine-Bold.ttf measures the exact same flat 600 — a monospace
// font's whole point is that weight never changes the advance grid — so
// there's no separate bold table to generate; this one table already
// answers both weights correctly.
const COURIER_WIDTH = 600;
const COURIER_WIDTHS = Object.fromEntries(
  Array.from({ length: 0x7e - 0x20 + 1 }, (_, i) => [String.fromCharCode(0x20 + i), COURIER_WIDTH]),
);

const FONT_TABLES = {
  arial: HELVETICA_WIDTHS,
  helvetica: HELVETICA_WIDTHS,
  'times new roman': TIMES_WIDTHS,
  times: TIMES_WIDTHS,
  calibri: CALIBRI_WIDTHS,
  georgia: GEORGIA_WIDTHS,
  courier: COURIER_WIDTHS,
  'courier new': COURIER_WIDTHS,
  consolas: COURIER_WIDTHS,
  monospace: COURIER_WIDTHS,
};

// A bold run draws as the real, wider bold-weight font (see
// HELVETICA_BOLD_WIDTHS' own comment) — consulted first whenever `bold` is
// true and the family has real bold data; falls through to FONT_TABLES'
// regular-weight entry for a family with no bold table of its own (or when
// `bold` is false). Monospace fonts have no entry here at all: COURIER_WIDTHS
// is already correct for both weights (see its own comment), so there's
// nothing a bold variant would change.
const BOLD_FONT_TABLES = {
  arial: HELVETICA_BOLD_WIDTHS,
  helvetica: HELVETICA_BOLD_WIDTHS,
  'times new roman': TIMES_BOLD_WIDTHS,
  times: TIMES_BOLD_WIDTHS,
  calibri: CALIBRI_BOLD_WIDTHS,
  georgia: GEORGIA_BOLD_WIDTHS,
};

// A character missing from a *named* table (e.g. an accented letter, an
// em-dash, an emoji) falls back to this flat average rather than 0 — close to
// the mean advance width across the tables above.
const DEFAULT_CHAR_WIDTH_EM = 0.52;

// None of the tables above have any entry beyond 0x7E — every one of them
// (and genericCharWidthEm's regex buckets) is Latin-script-shaped, so a
// non-Latin character always fell through to the flat ~0.5em default
// regardless of script. That's a severe underestimate for the scripts real
// documents actually mix in: a CJK/Hangul/fullwidth glyph is drawn in a near-
// square advance box (close to 1em, roughly double the Latin default), and a
// color emoji glyph is rendered at a similarly fixed near-square advance
// (often slightly wider still). A paragraph mixing Latin with even a few
// words of Chinese/Japanese/Korean or a couple of emoji therefore wrapped to
// visibly fewer lines than it actually renders to — undersizing the box
// reserved for it and letting the real text overflow into whatever shape was
// positioned right after it. This is checked ahead of the named-table/generic
// lookup (below) precisely because a declared Latin font family (Calibri,
// Arial, ...) has no bearing on how wide these glyphs draw — the renderer
// substitutes an actual CJK/emoji font for them regardless of what the run
// itself asked for.
const WIDE_SCRIPT_RANGES = [
  // Hiragana, Katakana, CJK Unified Ideographs (+ Extension A), Hangul
  // syllables, and CJK/ASCII fullwidth forms — all drawn full-width.
  { min: 0x3040, max: 0x30ff, widthEm: 1.0 },
  { min: 0x3400, max: 0x4dbf, widthEm: 1.0 },
  { min: 0x4e00, max: 0x9fff, widthEm: 1.0 },
  { min: 0xac00, max: 0xd7a3, widthEm: 1.0 },
  { min: 0xf900, max: 0xfaff, widthEm: 1.0 },
  { min: 0xff00, max: 0xffef, widthEm: 1.0 },
  // Indic scripts (Devanagari through Malayalam) and Thai — glyph clusters
  // (base + combining marks) run wide per source code point even though no
  // single codepoint is a full square the way a CJK ideograph is.
  { min: 0x0900, max: 0x0d7f, widthEm: 0.8 },
  { min: 0x0e00, max: 0x0e7f, widthEm: 0.8 },
  // Arabic (+ Arabic Supplement/Extended-A) and Cyrillic — narrower than the
  // scripts above but still wider on average than the ~0.52em Latin default.
  { min: 0x0600, max: 0x06ff, widthEm: 0.6 },
  { min: 0x0750, max: 0x077f, widthEm: 0.6 },
  { min: 0x08a0, max: 0x08ff, widthEm: 0.6 },
  { min: 0x0400, max: 0x04ff, widthEm: 0.62 },
  // Emoji/pictograph blocks (regional-indicator flags, misc symbols &
  // pictographs, dingbats, supplemental symbols & pictographs) — color glyphs
  // rendered at a fixed near-square (often slightly-over-1em) advance.
  { min: 0x1f1e6, max: 0x1f1ff, widthEm: 1.15 },
  { min: 0x2600, max: 0x27bf, widthEm: 1.15 },
  { min: 0x1f300, max: 0x1faff, widthEm: 1.15 },
];

function wideScriptCharWidthEm(ch) {
  const cp = ch.codePointAt(0);
  for (const { min, max, widthEm } of WIDE_SCRIPT_RANGES) {
    if (cp >= min && cp <= max) return widthEm;
  }
  return undefined;
}

// An unrecognized font-family (a custom web font, a font stack this module
// doesn't know) has no per-character table at all — bucket by character
// class instead of a single flat average, since a run of "iiii" and a run of
// "WWWW" at the same point size are nowhere close in real width, and
// collapsing both to one average measurably hurts short-text estimates
// (exactly the badge/pill case this module exists for).
const NARROW_CHARS = /[iIl.,;:'|!]/;
const WIDE_CHARS = /[mMWw@]/;
const UPPER_CHARS = /[A-Z]/;
const DIGIT_CHARS = /[0-9]/;

// Real bold tables above widen the average glyph by roughly 6-10% over its
// regular weight (e.g. HELVETICA_WIDTHS b:556 vs HELVETICA_BOLD_WIDTHS
// b:611) — applied here too so an unnamed/unrecognized font family (which
// only ever measures through this generic bucket, for either weight) gets
// the same bold correction a named family now gets from its own real table,
// instead of under-estimating a bold run's true width regardless of family.
const GENERIC_BOLD_WIDTH_MULTIPLIER = 1.08;

function genericCharWidthEm(ch, bold) {
  const base = ch === ' ' ? 0.27
    : NARROW_CHARS.test(ch) ? 0.28
    : WIDE_CHARS.test(ch) ? 0.85
    : UPPER_CHARS.test(ch) ? 0.68
    : DIGIT_CHARS.test(ch) ? 0.55
    : 0.5; // lowercase and anything else unclassified
  return bold ? base * GENERIC_BOLD_WIDTH_MULTIPLIER : base;
}

// A CSS font-family value may be a full stack ("Georgia, 'Times New Roman',
// serif") — only the first, most-specific entry is what actually renders, so
// only it is looked up; the rest (generic fallback keywords included) are
// ignored the same way a browser only falls through to them when the first
// font fails to load at all, which isn't a case this module can detect.
export function normalizeFontFamily(fontFamily) {
  if (!fontFamily) return undefined;
  const first = fontFamily.split(',')[0];
  return first.trim().replace(/^['"]|['"]$/g, '').toLowerCase();
}

function resolveFontTable(fontFamily, bold) {
  const key = normalizeFontFamily(fontFamily);
  if (!key) return undefined;
  if (bold && BOLD_FONT_TABLES[key]) return BOLD_FONT_TABLES[key];
  return FONT_TABLES[key];
}

function charWidthEm(ch, table) {
  return table[ch] != null ? table[ch] / 1000 : DEFAULT_CHAR_WIDTH_EM;
}

/**
 * Estimated rendered width of `text` at `fontSizePt`, in points.
 * `fontFamily` may be undefined or an unrecognized name — falls back to a
 * character-class-based generic estimate rather than a single flat average.
 * `bold`, when true, consults the family's own bold-weight table (or widens
 * the generic estimate — see genericCharWidthEm) instead of silently reusing
 * regular-weight metrics for a run that actually draws wider.
 */
export function estimateTextWidth(text, fontSizePt, fontFamily, bold = false) {
  const table = resolveFontTable(fontFamily, bold);
  let widthEm = 0;
  for (const ch of text) {
    widthEm += wideScriptCharWidthEm(ch) ?? (table ? charWidthEm(ch, table) : genericCharWidthEm(ch, bold));
  }
  return widthEm * fontSizePt;
}

/**
 * Greedily wraps `text` into lines no wider than `maxWidthPt`. A single word
 * wider than `maxWidthPt` on its own is placed alone on its own line rather
 * than looping forever trying to shrink it — there's no hyphenation here.
 *
 * `text` may itself carry literal `\n`s — flattenInline (idm/flatten-text.js)
 * emits one for every source `<br>` — which must stay a forced line break of
 * its own rather than being wrapped away: `\s` in the old single regex split
 * below matches `\n` too, so "Invoice No\nDate\nDue Date" (3 real lines) was
 * silently rejoined into one continuous word stream and re-wrapped purely by
 * width, collapsing down to as few as 1 line whenever the words happened to
 * fit the box's own width. Every caller here estimates a shape's height from
 * this function's own line count (see pptx/shapes.js's estimateBlockLineCount
 * and estimateShapeSize below) — an undercounted line count means an
 * undersized shape, which is what let a 3-line/6-line pair of paragraphs
 * overlap the very next shape on the slide. Splitting on `\n` first, then
 * width-wrapping each resulting segment independently, keeps every forced
 * break as its own line while still wrapping normally within each one.
 */
export function estimateWrappedLines(text, maxWidthPt, fontSizePt, fontFamily, bold = false) {
  const lines = [];
  for (const segment of (text ?? '').split('\n')) {
    lines.push(...wrapSegmentByWidth(segment, maxWidthPt, fontSizePt, fontFamily, bold));
  }
  return lines;
}

function wrapSegmentByWidth(text, maxWidthPt, fontSizePt, fontFamily, bold) {
  const words = text.split(/\s+/).filter(Boolean);
  if (words.length === 0) return [''];

  const spaceWidth = estimateTextWidth(' ', fontSizePt, fontFamily, bold);
  const lines = [];
  let current = '';
  let currentWidth = 0;

  for (const word of words) {
    const wordWidth = estimateTextWidth(word, fontSizePt, fontFamily, bold);
    if (current === '') {
      current = word;
      currentWidth = wordWidth;
      continue;
    }
    const withWord = currentWidth + spaceWidth + wordWidth;
    if (withWord <= maxWidthPt) {
      current += ' ' + word;
      currentWidth = withWord;
    } else {
      lines.push(current);
      current = word;
      currentWidth = wordWidth;
    }
  }
  lines.push(current);
  return lines;
}

/**
 * Estimates the shape size needed to hold `text` at `fontSizePt`, in points.
 *
 * - `options.maxWidthPt` set (a fixed-width box, e.g. a card column): wraps
 *   to that width and sizes height to the resulting line count.
 * - `options.maxWidthPt` omitted (an auto-sized badge/pill): sizes to one
 *   line's own width, biased up by `autoSizeBiasPercent` — the per-character
 *   tables are an estimate, not a measurement, so this bakes in a safety
 *   margin against clipping rather than fitting exactly to the guess.
 *
 * Returns `{ width, height, lineCount }` in points, both already inclusive
 * of `options.padding`.
 */
export function estimateShapeSize(text, fontSizePt, fontFamily, options = {}) {
  const {
    maxWidthPt,
    padding = {},
    lineHeightMultiplier = 1.2,
    autoSizeBiasPercent = 8,
  } = options;
  const padTop = padding.top ?? 0;
  const padRight = padding.right ?? 0;
  const padBottom = padding.bottom ?? 0;
  const padLeft = padding.left ?? 0;

  if (maxWidthPt != null) {
    const contentWidthPt = Math.max(1, maxWidthPt - padLeft - padRight);
    const lines = estimateWrappedLines(text, contentWidthPt, fontSizePt, fontFamily);
    const height = lines.length * fontSizePt * lineHeightMultiplier + padTop + padBottom;
    return { width: maxWidthPt, height, lineCount: lines.length };
  }

  const rawWidth = estimateTextWidth(text, fontSizePt, fontFamily);
  const width = rawWidth * (1 + autoSizeBiasPercent / 100) + padLeft + padRight;
  const height = fontSizePt * lineHeightMultiplier + padTop + padBottom;
  return { width, height, lineCount: 1 };
}
