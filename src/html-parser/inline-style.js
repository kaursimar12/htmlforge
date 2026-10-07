/**
 * Inline `style="..."` attribute parsing — a deliberately bounded, explicit
 * set of CSS properties that translate 1:1 into formatting concepts every
 * exporter already understands (bold/italic/underline/strikethrough/color/
 * background/font-size/text-align). This is NOT CSS cascade resolution and
 * NOT inferring semantic meaning from styling (e.g. "this looks like a
 * label") — it only carries forward an explicit declaration on the element
 * itself, the same way <strong>/<em> carry forward an explicit tag.
 */

// The full CSS Color Module Level 4 "extended keywords" set (147 names) —
// a small hand-picked list previously covered only the most common dozen,
// silently falling back to black for anything else (e.g. `border:5px double
// crimson` rendered a plain black border, with no warning that the color
// was ever lost).
const NAMED_COLORS = {
  aliceblue: 'F0F8FF', antiquewhite: 'FAEBD7', aqua: '00FFFF', aquamarine: '7FFFD4',
  azure: 'F0FFFF', beige: 'F5F5DC', bisque: 'FFE4C4', black: '000000',
  blanchedalmond: 'FFEBCD', blue: '0000FF', blueviolet: '8A2BE2', brown: 'A52A2A',
  burlywood: 'DEB887', cadetblue: '5F9EA0', chartreuse: '7FFF00', chocolate: 'D2691E',
  coral: 'FF7F50', cornflowerblue: '6495ED', cornsilk: 'FFF8DC', crimson: 'DC143C',
  cyan: '00FFFF', darkblue: '00008B', darkcyan: '008B8B', darkgoldenrod: 'B8860B',
  darkgray: 'A9A9A9', darkgreen: '006400', darkgrey: 'A9A9A9', darkkhaki: 'BDB76B',
  darkmagenta: '8B008B', darkolivegreen: '556B2F', darkorange: 'FF8C00', darkorchid: '9932CC',
  darkred: '8B0000', darksalmon: 'E9967A', darkseagreen: '8FBC8F', darkslateblue: '483D8B',
  darkslategray: '2F4F4F', darkslategrey: '2F4F4F', darkturquoise: '00CED1', darkviolet: '9400D3',
  deeppink: 'FF1493', deepskyblue: '00BFFF', dimgray: '696969', dimgrey: '696969',
  dodgerblue: '1E90FF', firebrick: 'B22222', floralwhite: 'FFFAF0', forestgreen: '228B22',
  fuchsia: 'FF00FF', gainsboro: 'DCDCDC', ghostwhite: 'F8F8FF', gold: 'FFD700',
  goldenrod: 'DAA520', gray: '808080', green: '008000', greenyellow: 'ADFF2F',
  grey: '808080', honeydew: 'F0FFF0', hotpink: 'FF69B4', indianred: 'CD5C5C',
  indigo: '4B0082', ivory: 'FFFFF0', khaki: 'F0E68C', lavender: 'E6E6FA',
  lavenderblush: 'FFF0F5', lawngreen: '7CFC00', lemonchiffon: 'FFFACD', lightblue: 'ADD8E6',
  lightcoral: 'F08080', lightcyan: 'E0FFFF', lightgoldenrodyellow: 'FAFAD2', lightgray: 'D3D3D3',
  lightgreen: '90EE90', lightgrey: 'D3D3D3', lightpink: 'FFB6C1', lightsalmon: 'FFA07A',
  lightseagreen: '20B2AA', lightskyblue: '87CEFA', lightslategray: '778899', lightslategrey: '778899',
  lightsteelblue: 'B0C4DE', lightyellow: 'FFFFE0', lime: '00FF00', limegreen: '32CD32',
  linen: 'FAF0E6', magenta: 'FF00FF', maroon: '800000', mediumaquamarine: '66CDAA',
  mediumblue: '0000CD', mediumorchid: 'BA55D3', mediumpurple: '9370DB', mediumseagreen: '3CB371',
  mediumslateblue: '7B68EE', mediumspringgreen: '00FA9A', mediumturquoise: '48D1CC', mediumvioletred: 'C71585',
  midnightblue: '191970', mintcream: 'F5FFFA', mistyrose: 'FFE4E1', moccasin: 'FFE4B5',
  navajowhite: 'FFDEAD', navy: '000080', oldlace: 'FDF5E6', olive: '808000',
  olivedrab: '6B8E23', orange: 'FFA500', orangered: 'FF4500', orchid: 'DA70D6',
  palegoldenrod: 'EEE8AA', palegreen: '98FB98', paleturquoise: 'AFEEEE', palevioletred: 'DB7093',
  papayawhip: 'FFEFD5', peachpuff: 'FFDAB9', peru: 'CD853F', pink: 'FFC0CB',
  plum: 'DDA0DD', powderblue: 'B0E0E6', purple: '800080', rebeccapurple: '663399',
  red: 'FF0000', rosybrown: 'BC8F8F', royalblue: '4169E1', saddlebrown: '8B4513',
  salmon: 'FA8072', sandybrown: 'F4A460', seagreen: '2E8B57', seashell: 'FFF5EE',
  sienna: 'A0522D', silver: 'C0C0C0', skyblue: '87CEEB', slateblue: '6A5ACD',
  slategray: '708090', slategrey: '708090', snow: 'FFFAFA', springgreen: '00FF7F',
  steelblue: '4682B4', tan: 'D2B48C', teal: '008080', thistle: 'D8BFD8',
  tomato: 'FF6347', turquoise: '40E0D0', violet: 'EE82EE', wheat: 'F5DEB3',
  white: 'FFFFFF', whitesmoke: 'F5F5F5', yellow: 'FFFF00', yellowgreen: '9ACD32',
};

const ALIGNMENTS = new Set(['left', 'center', 'right', 'justify']);

const BORDER_STYLES = new Set(['solid', 'dashed', 'dotted', 'double', 'none', 'hidden']);

// Every extractXYZ function below independently calls parseInlineStyleAttr
// on the *same* element's style string (mapCells alone calls ~10 of them per
// <td>), and separately, many elements share the exact same resolved style
// string in the first place (e.g. every <th> matched by the same `thead
// th{...}` rule via applyClassStyles) — so the identical string gets
// re-split from scratch far more often than the actual distinct-string count
// warrants. Caching by the raw string, once, removes that redundant work
// without changing any extractor's signature or return value: same string
// in, same object out, every time, whether computed fresh or served from
// the cache — a pure function's result never depends on being reused.
const parsedStyleCache = new Map();

// The CSS `font` shorthand (`font: [style] [weight] size[/line-height]
// family`) is an extremely common way to set body typography in one
// declaration (e.g. `body{font:16px/1.6 Calibri,Arial,sans-serif}`) — far
// more common in real-world pages than spelling out `font-size`/
// `line-height`/`font-family` as separate longhands. Without expanding it
// here, every longhand-only reader below (extractRunStyle's `font-size`,
// this file's own line-height extractor) sees nothing at all for a page
// that only ever uses the shorthand, silently falling back to this
// project's own hardcoded defaults instead of the page's actual base font.
// Only the common numeric-size-plus-family form is matched — CSS-wide
// system-font keywords (`caption`, `menu`, `icon`, ...) and font-variant are
// out of scope, same "common real-world pattern, not full CSS parsing"
// bound this file already accepts for every other shorthand it expands.
const FONT_SHORTHAND = /^(?:(italic|oblique)\s+)?(?:(bold|bolder|lighter|\d{3,4})\s+)?([\d.]+(?:px|pt))(?:\s*\/\s*([\d.]+))?\s+(.+)$/i;

function expandFontShorthand(value) {
  const match = FONT_SHORTHAND.exec(value.trim());
  if (!match) return undefined;
  const [, style, weight, size, lineHeight, family] = match;
  const out = { 'font-size': size, 'font-family': family };
  if (style) out['font-style'] = style;
  if (weight) out['font-weight'] = weight;
  if (lineHeight) out['line-height'] = lineHeight;
  return out;
}

export function parseInlineStyleAttr(styleAttr) {
  const key = styleAttr ?? '';
  const cached = parsedStyleCache.get(key);
  if (cached) return cached;

  const out = {};
  for (const decl of key.split(';')) {
    const i = decl.indexOf(':');
    if (i === -1) continue;
    const prop = decl.slice(0, i).trim().toLowerCase();
    const value = decl.slice(i + 1).trim();
    if (!prop || !value) continue;
    // The shorthand's own sub-values are set first, at the shorthand's own
    // position in declaration order, so a longhand declared LATER in the
    // same style string still correctly overrides it (real CSS's own
    // "later declaration wins" rule) — only a longhand declared earlier
    // gets clobbered, matching a real shorthand's own reset behavior.
    if (prop === 'font') {
      const expanded = expandFontShorthand(value);
      if (expanded) { Object.assign(out, expanded); continue; }
    }
    out[prop] = value;
  }
  parsedStyleCache.set(key, out);
  return out;
}

/** Character-level formatting (bold/italic/underline/strikethrough/color/background/size). */
export function extractRunStyle(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const style = {};

  const color = normalizeColor(css['color']);
  if (color) style.color = color;

  const background = resolveBackgroundColor(css);
  if (background) style.backgroundColor = background;

  const weight = css['font-weight'];
  if (weight === 'bold' || weight === 'bolder' || (/^\d+$/.test(weight ?? '') && Number(weight) >= 600)) {
    style.bold = true;
  }
  // The boolean above is all DOCX/PPTX/PDF's own bold-run-property readers
  // ever needed (a run is either the bold face or it isn't). A font
  // resolver picking a specific weight's file out of a multi-weight
  // @font-face family (e.g. Poppins 400 vs 700) needs the actual number,
  // not a >=600 threshold — "bold: true" can't tell a 600 from an 800.
  // `bolder`/`lighter` are relative-to-inherited-weight keywords with no
  // inheritance tracking in this project (same documented bound as the rest
  // of this file); approximated as the nearest common absolute weight
  // (700/300) rather than left unresolved. Carried alongside `bold`, never
  // replacing it — every existing boolean consumer is unaffected.
  const numericWeight = weight === 'bold' ? 700
    : weight === 'bolder' ? 700
    : weight === 'lighter' ? 300
    : weight === 'normal' ? 400
    : /^\d{3}$/.test(weight ?? '') ? Number(weight)
    : undefined;
  if (numericWeight !== undefined) style.fontWeight = numericWeight;

  if (css['font-style'] === 'italic' || css['font-style'] === 'oblique') style.italic = true;

  const decoration = css['text-decoration'] ?? css['text-decoration-line'];
  if (decoration) {
    if (decoration.includes('underline')) style.underline = true;
    if (decoration.includes('line-through')) style.strikethrough = true;
  }

  const fontSize = normalizeFontSizePt(css['font-size']);
  if (fontSize) style.fontSizePt = fontSize;

  const fontFamily = extractFontFamily(css['font-family']);
  if (fontFamily) style.fontFamily = fontFamily;

  return Object.keys(style).length ? style : undefined;
}

// font-family is a comma-separated fallback stack (e.g. "'Times New Roman',
// serif") — only the first name is used, matching how a renderer picks the
// first available font from the stack; surrounding quotes are stripped. The
// name is passed through as-is (not validated against installed fonts) —
// the target renderer resolves or substitutes it, the same as a browser does.
function extractFontFamily(value) {
  if (!value) return undefined;
  const first = value.split(',')[0].trim().replace(/^['"]|['"]$/g, '');
  return first || undefined;
}

// margin's 1/2/3/4-value shorthand (CSS §8.3) — only the vertical (top/
// bottom) components are extracted, since OOXML paragraph spacing has no
// horizontal-margin equivalent (that would be indentation, a different,
// separately-scoped concept not modeled here).
function parseMarginShorthand(value) {
  const parts = (value ?? '').trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return {};
  if (parts.length === 1) return { top: parts[0], bottom: parts[0] };
  if (parts.length === 2) return { top: parts[0], bottom: parts[0] };
  return { top: parts[0], bottom: parts[2] }; // 3- or 4-value form: top ... bottom ...
}

/**
 * An element's own vertical margin (`margin-top`/`margin-bottom`, or the
 * `margin` shorthand) — the paragraph-spacing equivalent of border/
 * background-color, carried the same way onto whichever paragraph/heading
 * the element produced (see dom-to-idm.js's mapElement). Returns undefined
 * when neither side is set to a resolvable absolute length.
 */
export function extractMarginPt(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const shorthand = parseMarginShorthand(css['margin']);
  const beforePt = parseAbsoluteLengthPt(css['margin-top'] ?? shorthand.top);
  const afterPt = parseAbsoluteLengthPt(css['margin-bottom'] ?? shorthand.bottom);
  if (beforePt === undefined && afterPt === undefined) return undefined;
  return { beforePt, afterPt };
}

// left/right sides of the margin shorthand — parseMarginShorthand above only
// tracks top/bottom (the only sides paragraph spacing cares about); this
// mirrors it for left/right, needed only to detect the
// `margin-left:auto`/`margin-right:auto`/`margin:0 auto` centering idiom.
function marginShorthandHorizontal(value) {
  const parts = (value ?? '').trim().split(/\s+/).filter(Boolean);
  if (parts.length === 1) return { left: parts[0], right: parts[0] };
  if (parts.length === 2 || parts.length === 3) return { left: parts[1], right: parts[1] };
  if (parts.length === 4) return { left: parts[3], right: parts[1] };
  return {};
}

/**
 * The `margin-left:auto`/`margin-right:auto` idiom — the standard way a
 * fixed/percentage-width block is pushed to one side or centered without
 * flex/grid (`table{width:50%;margin-left:auto}` hugs the right edge,
 * `margin:0 auto` centers). Only consumed by the XLSX exporter today, to
 * decide which sheet columns a table should occupy relative to the
 * document's full column span (see dom-to-idm.js's mapTable).
 */
export function extractMarginAuto(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const shorthand = marginShorthandHorizontal(css['margin']);
  const left = (css['margin-left'] ?? shorthand.left ?? '').trim().toLowerCase();
  const right = (css['margin-right'] ?? shorthand.right ?? '').trim().toLowerCase();
  if (left === 'auto' && right === 'auto') return 'both';
  if (left === 'auto') return 'left';
  if (right === 'auto') return 'right';
  return undefined;
}

// padding's 1/2/3/4-value shorthand (CSS §8.4) — same expansion rule as margin.
function parsePaddingShorthand(value) {
  const parts = (value ?? '').trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return {};
  if (parts.length === 1) return { top: parts[0], right: parts[0], bottom: parts[0], left: parts[0] };
  if (parts.length === 2) return { top: parts[0], right: parts[1], bottom: parts[0], left: parts[1] };
  if (parts.length === 3) return { top: parts[0], right: parts[1], bottom: parts[2], left: parts[1] };
  return { top: parts[0], right: parts[1], bottom: parts[2], left: parts[3] };
}

/**
 * An element's own padding (`padding-top/right/bottom/left`, or the
 * `padding` shorthand) — the space *inside* a border/background, as
 * opposed to extractMarginPt's space *outside* it. Maps to a bordered
 * paragraph's border-to-text spacing (OOXML's w:pBdr `w:space` — one of the
 * few paragraph properties that has a direct "space inside a border"
 * concept) or, for a background-only box with no border, an approximate
 * horizontal indent (see dom-to-idm.js's mapElement / document.js's
 * renderPBdr/renderIndent). Returns undefined when no side resolves to an
 * absolute length.
 */
export function extractPaddingPt(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const shorthand = parsePaddingShorthand(css['padding']);
  const top = parseAbsoluteLengthPt(css['padding-top'] ?? shorthand.top);
  const right = parseAbsoluteLengthPt(css['padding-right'] ?? shorthand.right);
  const bottom = parseAbsoluteLengthPt(css['padding-bottom'] ?? shorthand.bottom);
  const left = parseAbsoluteLengthPt(css['padding-left'] ?? shorthand.left);
  if ([top, right, bottom, left].every((v) => v === undefined)) return undefined;
  return { top, right, bottom, left };
}

/**
 * A unitless `line-height` (e.g. `1.6`, from `line-height:1.6` or the `font`
 * shorthand's `size/line-height` part) is a MULTIPLIER of an element's own
 * font-size, inherited by descendants as that same multiplier rather than a
 * fixed length — the standard way body typography sets line spacing. Only
 * the unitless numeric form is modeled (by far the most common way it's
 * declared); a length-valued line-height (`24px`) would need dividing by
 * the element's own resolved font-size to get an equivalent multiplier,
 * which this project doesn't track independently, so it's left unmatched,
 * same as any other unsupported value elsewhere in this file.
 */
export function extractLineHeight(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const value = css['line-height'];
  if (!value) return undefined;
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? n : undefined;
}

/** Paragraph-level formatting (currently just text-align). */
export function extractAlign(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const align = css['text-align']?.toLowerCase();
  return align && ALIGNMENTS.has(align) ? align : undefined;
}

/**
 * A container's `display:flex`/`display:grid` — the signal dom-to-idm.js's
 * mapLayoutBlocks uses to lay its children out as real side-by-side/wrapped
 * cells (PPTX) instead of the plain vertical text-flow every `<div>`
 * otherwise gets. Any other `display` value (block/inline-block/none/...)
 * has no dedicated layout handling and returns undefined, same as if the
 * property were absent.
 */
export function extractDisplayMode(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const display = css['display']?.toLowerCase();
  return display === 'flex' || display === 'grid' ? display : undefined;
}

/**
 * `grid-template-columns`'s track count — only the count matters here (not
 * each track's own size), since there's no real layout engine to honor
 * unequal track sizes anyway (see pptx/slide.js's renderLayoutBlock, which
 * always splits the row into equal-width columns). Handles the plain
 * space-separated form (`1fr 1fr 1fr` -> 3) and `repeat(N, ...)` (`repeat(3,
 * 1fr)` -> 3); anything else unparseable returns undefined and the caller
 * falls back to treating every cell as its own column (a single row).
 */
export function extractGridColumnCount(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const value = css['grid-template-columns'];
  if (!value) return undefined;
  const repeatMatch = /^repeat\(\s*(\d+)\s*,/i.exec(value.trim());
  if (repeatMatch) return Number(repeatMatch[1]);
  const count = value.trim().split(/\s+/).filter(Boolean).length;
  return count || undefined;
}

/**
 * A flex/grid container's `gap` (or the `column-gap`/`row-gap` longhands) —
 * only the first resolvable value is used as one uniform spacing in both
 * directions, same "no real layout engine" simplification extractGridColumnCount
 * documents. Returns undefined (caller falls back to its own default) when
 * unset or unparseable.
 */
export function extractGapPt(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const value = css['gap'] ?? css['column-gap'] ?? css['row-gap'];
  return parseAbsoluteLengthPt(value?.trim().split(/\s+/)[0]);
}

// CSS `vertical-align`'s box-alignment values (top/middle/bottom) map
// directly to a table cell's vertical position — 'middle' is renamed to
// 'center' to match this project's own paragraph/cell alignment vocabulary
// (and XLSX's own `<alignment vertical="...">` attribute values). The
// property's *other* values (baseline/sub/super/text-top/text-bottom, which
// only make sense for inline-level boxes, not a table cell) have no
// equivalent here and are left unmapped, same as an unrecognized
// `text-align` value.
const VERTICAL_ALIGNS = { top: 'top', middle: 'center', bottom: 'bottom' };

export function extractVerticalAlign(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const value = css['vertical-align']?.toLowerCase();
  return value ? VERTICAL_ALIGNS[value] : undefined;
}

// `display:none` and `visibility:hidden` are treated identically here: no
// target format has a "present in the document but not shown" concept for
// arbitrary content (Word's `w:vanish` comes close for a text run, but
// nothing analogous exists for a block/cell, and Excel has nothing at all),
// so the only consistent, format-agnostic mapping is "the author didn't want
// this seen, so it doesn't appear in the exported document either" — the
// same treatment already given to <script>/<style>/<nav> noise elements.
export function isHidden(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  return css['display']?.toLowerCase() === 'none' || css['visibility']?.toLowerCase() === 'hidden';
}

// A real (non-`normal`) `white-space` — the CSS-level signal that a plain
// `<div>`/`<span>` styled to *look* like a code block (dark background,
// monospace font, `white-space:pre`) should preserve its literal line
// breaks/blank lines the exact same way a semantic `<pre>` tag already does,
// rather than collapsing them the ordinary HTML way — a common real-world
// pattern from CMS/editor output that doesn't bother with the `<pre>` tag
// itself. `pre-wrap`/`pre-line` both still preserve line breaks (they only
// differ from `pre` in whether long lines additionally wrap), so all three
// map the same way here; `nowrap` doesn't preserve breaks at all and is
// deliberately excluded.
export function preservesWhitespace(styleAttr) {
  const value = parseInlineStyleAttr(styleAttr)['white-space']?.toLowerCase();
  return value === 'pre' || value === 'pre-wrap' || value === 'pre-line';
}

// Only 'rtl' is worth carrying — 'ltr' is the default in every target format,
// so leaving it unset already means the same thing.
export function extractDirection(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  return css['direction']?.toLowerCase() === 'rtl' ? 'rtl' : undefined;
}

// Only the top-to-bottom variant is mapped — it's the one both DOCX
// (`w:textDirection val="tbRl"`) and XLSX (`textRotation="255"`) have a
// direct, well-defined equivalent for; 'vertical-lr'/'sideways-lr' have no
// equivalently clean mapping and are left unrecognized.
export function extractWritingMode(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  return css['writing-mode']?.toLowerCase() === 'vertical-rl' ? 'vertical-rl' : undefined;
}

// CSS `transform: rotate(Ndeg)` — only this single-function, degree-unit form
// is recognized (no combined transforms, no turn/rad/grad units); positive N
// is clockwise, matching CSS's own convention. Anything else (multiple
// transforms, non-rotate functions, other angle units) is left unset rather
// than guessed at.
export function extractRotationDeg(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const match = /^rotate\(\s*(-?[\d.]+)deg\s*\)$/.exec((css['transform'] ?? '').trim());
  return match ? Number(match[1]) : undefined;
}

/**
 * A <td>/<tr>'s explicit pixel height (e.g. a decorative colored strip cell
 * styled `height:4px`) — percentage/em/other units aren't handled since they
 * don't translate to an absolute twip value the way px does.
 */
export function extractHeightPx(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const match = /^([\d.]+)px$/.exec((css['height'] ?? '').trim());
  return match ? Number(match[1]) : undefined;
}

/**
 * An element's own explicit `max-height` in pixels — same px-only scope as
 * extractHeightPx. Read alongside extractHeightPx/extractOverflow* by
 * dom-to-idm.js's scrollable-container detection: a declared height OR
 * max-height, combined with overflow:auto/scroll/hidden, is what marks a
 * `<div>` as a "scrollable container" for the preserveScrollableContent
 * option (see resolve-scrollable-content.js).
 */
export function extractMaxHeightPx(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const match = /^([\d.]+)px$/.exec((css['max-height'] ?? '').trim());
  return match ? Number(match[1]) : undefined;
}

const OVERFLOW_VALUES = new Set(['visible', 'hidden', 'scroll', 'auto']);

function normalizeOverflowValue(value) {
  const lower = (value ?? '').trim().toLowerCase();
  return OVERFLOW_VALUES.has(lower) ? lower : undefined;
}

/**
 * CSS `overflow` (shorthand for both axes) — one of 'visible'/'hidden'/
 * 'scroll'/'auto', or undefined if unset/unrecognized. `overflow-x`/
 * `overflow-y` (below) take precedence when present, matching CSS's own
 * shorthand-vs-longhand cascade.
 */
export function extractOverflow(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  return normalizeOverflowValue(css['overflow']);
}

export function extractOverflowX(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  return normalizeOverflowValue(css['overflow-x']) ?? normalizeOverflowValue(css['overflow']);
}

export function extractOverflowY(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  return normalizeOverflowValue(css['overflow-y']) ?? normalizeOverflowValue(css['overflow']);
}

/**
 * An element's own explicit pixel width — e.g. `.circle{width:120px}`. Used
 * (alongside isInlineBlock) by pptx/shapes.js as the real "this box should
 * shrink to a specific size, not stretch to fill its container" signal,
 * instead of guessing shrink-vs-stretch from how short the text happens to
 * be. Percentage/em/other units aren't handled, same as extractHeightPx.
 */
export function extractWidthPx(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const match = /^([\d.]+)px$/.exec((css['width'] ?? '').trim());
  return match ? Number(match[1]) : undefined;
}

/**
 * An element's own explicit percentage width — e.g. `table{width:50%}`. A
 * proportion of the containing block rather than an absolute size, so it's
 * kept separate from extractWidthPx. Only consumed today by the XLSX
 * exporter (dom-to-idm.js's mapTable) to decide how many of the sheet's
 * columns a table should span. Named distinctly from dom-to-idm.js's own
 * (unrelated, cell-level) `extractWidthPercent` — that one reads a `<td
 * width="52%">` HTML attribute or style for a table CELL's proportional
 * column width; this one reads only the CSS `width` property, for a whole
 * TABLE's width relative to its container.
 */
export function extractTableWidthPercent(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const match = /^([\d.]+)%$/.exec((css['width'] ?? '').trim());
  return match ? Number(match[1]) : undefined;
}

// `body{max-width:1000px;margin:auto}` is the standard "centered page-like
// container" pattern (used everywhere from Bootstrap's `.container` to any
// hand-rolled page template) — it's the actual visual width a browser
// renders the page at, independent of the viewport. See dom-to-idm.js's own
// use of this for sizing a PDF page to match that width instead of a fixed
// Letter/A4 default. Percentage/em/other units aren't handled, same as
// extractWidthPx's own scope.
export function extractMaxWidthPx(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const match = /^([\d.]+)px$/.exec((css['max-width'] ?? '').trim());
  return match ? Number(match[1]) : undefined;
}

/**
 * Whether an element declares `display:inline-block` — the other "shrink to
 * content, don't stretch" signal alongside an explicit width (see
 * extractWidthPx): the CSS property that actually distinguishes a
 * shrink-wrapped pill/badge (`.badge{display:inline-block;padding:6px
 * 12px;border-radius:20px}`) from an ordinary block-level `<div>` (which
 * defaults to filling its container's width — e.g. a full-width gradient
 * banner/button with the exact same short-text-plus-background shape but no
 * inline-block declaration).
 */
export function isInlineBlock(styleAttr) {
  return parseInlineStyleAttr(styleAttr)['display']?.toLowerCase() === 'inline-block';
}

/**
 * Whether an element declares a real (non-`none`) `box-shadow` — DrawingML
 * has no way to reproduce a CSS shadow's exact offset/blur/spread/color, so
 * this only ever answers yes/no; pptx/shapes.js's shadowEffectXml applies
 * one fixed, PowerPoint-native-looking drop shadow regardless of the CSS
 * value's own parameters, the same "presence, not precision" approximation
 * extractBackgroundGradient's own angle/stops (an actually representable
 * DrawingML construct) is a more faithful version of.
 */
export function hasBoxShadow(styleAttr) {
  const value = parseInlineStyleAttr(styleAttr)['box-shadow'];
  return !!value && value.trim().toLowerCase() !== 'none';
}

/**
 * Whether an element declares a real (non-zero) `border-radius` — OOXML
 * paragraph borders and table shading have no rounded-corner concept at all
 * (see extractTableBorder's own note), so like hasBoxShadow this only ever
 * answers yes/no. Used by pptx/shapes.js and docx/document.js purely to
 * decide `roundRect` vs `rect` shape geometry when promoting a bordered/
 * backgrounded box to a real DrawingML shape — independent of whether that
 * box also has an actual `border` (a borderless gradient panel with its own
 * border-radius still needs `roundRect`).
 */
export function hasBorderRadius(styleAttr) {
  const value = parseInlineStyleAttr(styleAttr)['border-radius'];
  if (!value) return false;
  return !/^0(\.0+)?(px|pt|%|em|rem)?$/.test(value.trim());
}

/**
 * Whether `border-radius` is declared large enough to turn the box into a
 * full ellipse (the standard CSS circle/oval idiom — `border-radius:50%`
 * on a square box is a circle, on a non-square box it's an oval that hugs
 * the box's own aspect ratio either way). Anything below 50% just rounds the
 * corners partially — see hasBorderRadius — and stays a `roundRect`, not an
 * `ellipse`.
 *
 * Only the percentage form is recognized: that's the real-world idiom this
 * targets. A large pixel radius (`border-radius:999px`) produces a
 * "stadium"/pill shape (rounded rectangle with fully-round ends) rather than
 * an ellipse on a non-square box — a different shape with no evidence of
 * real-world use in this project's target HTML sources, so it's left alone
 * rather than guessed at.
 */
export function isCircularBorderRadius(styleAttr) {
  const value = parseInlineStyleAttr(styleAttr)['border-radius'];
  if (!value) return false;
  const match = /^([\d.]+)%$/.exec(value.trim());
  return !!match && Number(match[1]) >= 50;
}

/**
 * A <td>/<th>'s own background-color (e.g. a decorative colored strip cell,
 * or a shaded header cell) — separate from extractRunStyle's backgroundColor,
 * which highlights the run's *text*. Same CSS property, different target
 * construct depending on which element carries it.
 */
export function extractBackgroundColor(styleAttr) {
  return resolveBackgroundColor(parseInlineStyleAttr(styleAttr));
}

// `background-color` takes priority; `background` (the shorthand — very
// common as just a bare color, e.g. style="background:yellow") is tried as
// a fallback. A plain color value resolves directly; a gradient function
// falls back to its first explicit color stop (see below). Anything else
// (e.g. "url(...) no-repeat") simply fails to normalize as a color and is
// left unset, same as any unrecognized value.
function resolveBackgroundColor(css) {
  const value = css['background-color'] ?? css['background'];
  return normalizeColor(value) ?? extractGradientFirstColor(value);
}

// A gradient has no flat-fill OOXML equivalent — approximating it with its
// first explicit color stop (a common fallback strategy, e.g. what an
// email client without gradient support does) is a deterministic use of a
// color the author actually wrote, not an invented one, so it's preferable
// to leaving the background entirely blank. The gradient's direction/shape
// argument (e.g. "90deg", "to right", "circle at center") sits in the same
// first comma-separated segment as a color could, so each segment's first
// word is tried in order until one resolves as a real color.
function extractGradientFirstColor(value) {
  if (!value) return undefined;
  const match = /(?:linear|radial|conic)-gradient\(([^)]*)\)/i.exec(value);
  if (!match) return undefined;
  for (const stop of match[1].split(',')) {
    const color = normalizeColor(stop.trim().split(/\s+/)[0]);
    if (color) return color;
  }
  return undefined;
}

// CSS angle keywords, converted to the equivalent explicit degree value —
// "to bottom" (CSS's own default with no angle at all) is the 180deg case.
const GRADIENT_KEYWORD_ANGLE = {
  'to top': 0, 'to right': 90, 'to bottom': 180, 'to left': 270,
  'to top right': 45, 'to right top': 45,
  'to bottom right': 135, 'to right bottom': 135,
  'to bottom left': 225, 'to left bottom': 225,
  'to top left': 315, 'to left top': 315,
};

/**
 * Unlike extractGradientFirstColor's single-color approximation, this keeps
 * every color stop and the gradient's direction — DrawingML (PPTX) has a
 * real `a:gradFill` construct, so a linear gradient background (the common
 * "gradient button/panel" pattern) doesn't need to degrade to a flat fill
 * there the way DOCX/XLSX still must. Returns `{ angleDeg, stops }` (stops
 * as `{ color, offset }`, offset 0-100) for a `linear-gradient(...)` with 2+
 * recognizable color stops, or undefined for anything else (radial/conic —
 * no direct DrawingML equivalent — a single-stop gradient, or an
 * unparseable value) so callers fall back to extractGradientFirstColor's
 * flat approximation.
 */
export function extractBackgroundGradient(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const value = css['background-color'] ?? css['background'];
  if (!value) return undefined;
  const match = /^linear-gradient\(([^)]*)\)$/i.exec(value.trim());
  if (!match) return undefined;

  const parts = match[1].split(',').map((s) => s.trim());
  let angleDeg = 180; // CSS default direction with no explicit angle: top-to-bottom
  let start = 0;
  const angleMatch = /^(-?[\d.]+)deg$/i.exec(parts[0]);
  if (angleMatch) {
    angleDeg = Number(angleMatch[1]);
    start = 1;
  } else if (GRADIENT_KEYWORD_ANGLE[parts[0].toLowerCase()] !== undefined) {
    angleDeg = GRADIENT_KEYWORD_ANGLE[parts[0].toLowerCase()];
    start = 1;
  }

  const colorParts = parts.slice(start);
  const stops = [];
  colorParts.forEach((part, i) => {
    const tokens = part.split(/\s+/);
    const color = normalizeColor(tokens[0]);
    if (!color) return;
    const pctMatch = /^([\d.]+)%$/.exec(tokens[1] ?? '');
    const offset = pctMatch ? Number(pctMatch[1])
      : colorParts.length > 1 ? Math.round((i / (colorParts.length - 1)) * 100) : 0;
    stops.push({ color, offset });
  });
  return stops.length >= 2 ? { angleDeg, stops } : undefined;
}

// Shared by extractTableBorder/extractCellBorder — pulls width/style/color
// tokens out of a `border[-side]` shorthand value (order-independent, as CSS
// allows). Longhand border[-side]-width/style/color properties, checked by
// the callers, always take priority over whatever the shorthand implied.
function parseBorderShorthand(value) {
  let width, color, style;
  for (const token of (value ?? '').trim().split(/\s+/).filter(Boolean)) {
    if (BORDER_STYLES.has(token.toLowerCase())) style = token.toLowerCase();
    else if (/^[\d.]+(px|pt)$/.test(token)) width = token;
    else if (normalizeColor(token)) color = token;
  }
  return { width, color, style };
}

// CSS border-style values that have a direct OOXML ST_Border equivalent.
// Anything else explicit (groove/ridge/inset/outset — rare, no clean OOXML
// match) or absent falls back to a plain single line, same as before this
// field existed.
const OOXML_BORDER_STYLE = { solid: 'single', dashed: 'dashed', dotted: 'dotted', double: 'double' };

function resolveBorder(width, color, style) {
  if (style === 'none' || style === 'hidden') return undefined;
  const resolvedColor = normalizeColor(color);
  if (!resolvedColor && !width) return undefined;
  // A keyword border-width ('thin'/'medium'/'thick') isn't a supported unit
  // here — same as an unrecognized value, it falls back to the hairline default.
  const widthPt = parseAbsoluteLengthPt(width) ?? 0.75;
  return {
    color: resolvedColor ?? '000000',
    sizeEighthPt: Math.max(2, Math.round(widthPt * 8)),
    style: OOXML_BORDER_STYLE[style] ?? 'single',
  };
}

/**
 * An element's own `border` (shorthand `border: 1px solid #hex` or the
 * longhand border-width/style/color) — despite the name, this reads the
 * generic CSS `border` property, not something table-specific; it's used
 * both for a <table>'s own border and, via dom-to-idm.js's mapElement, a
 * plain block element's (e.g. <div>) border. Returns undefined when no
 * border is declared or it's explicitly none/hidden — callers fall back to
 * their own default border behavior in that case. `border-radius` has no
 * OOXML border equivalent and is intentionally not modeled.
 */
export function extractTableBorder(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const shorthand = parseBorderShorthand(css['border']);
  const width = css['border-width'] ?? shorthand.width;
  const color = css['border-color'] ?? shorthand.color;
  const style = css['border-style']?.toLowerCase() ?? shorthand.style;
  return resolveBorder(width, color, style);
}

/**
 * Whether a <td>/<th> explicitly cancels its border via `border:none`/
 * `border-style:none` (or `hidden`) — as opposed to simply not declaring any
 * border at all. `resolveBorder` returns `undefined` for both cases, which
 * is fine for a plain paragraph/div (nothing to override), but a table cell
 * can inherit a default grid border from the *table's own style* (e.g. a
 * data table's TableGrid) when it has no border info of its own — the
 * common real-world pattern being a single giant `<table>` used for overall
 * page layout, with a base `td{border:1px solid #ccc}` tag rule and a
 * `.no-border{border:none}` class applied to the purely-positional cells.
 * Those cells need to actively suppress the table style's default border
 * (OOXML's `w:val="nil"`), not just contribute no border info of their own.
 */
export function hasExplicitNoBorder(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  const shorthand = parseBorderShorthand(css['border']);
  const style = css['border-style']?.toLowerCase() ?? shorthand.style;
  return style === 'none' || style === 'hidden';
}

// Shared by extractCellBorder/extractBlockSideBorder. `fallbackToGeneric`
// controls whether an element's plain `border`/`border-width`/`border-color`/
// `border-style` shorthand (applied uniformly to all four sides) is read as
// a fallback when no side-specific declaration exists for `side` — a table
// cell has no separate "uniform border" concept of its own, so it needs that
// fallback; a block element already gets its uniform border from
// extractTableBorder, so extractBlockSideBorder deliberately omits it to
// avoid reporting the same border twice under two different fields.
function extractSideBorder(styleAttr, side, fallbackToGeneric) {
  const css = parseInlineStyleAttr(styleAttr);
  const sideShorthand = parseBorderShorthand(css[`border-${side}`]);
  const allShorthand = fallbackToGeneric ? parseBorderShorthand(css['border']) : {};
  const width = css[`border-${side}-width`] ?? sideShorthand.width ?? (fallbackToGeneric ? css['border-width'] ?? allShorthand.width : undefined);
  const color = css[`border-${side}-color`] ?? sideShorthand.color ?? (fallbackToGeneric ? css['border-color'] ?? allShorthand.color : undefined);
  const style = css[`border-${side}-style`]?.toLowerCase() ?? sideShorthand.style ?? (fallbackToGeneric ? css['border-style']?.toLowerCase() ?? allShorthand.style : undefined);
  return resolveBorder(width, color, style);
}

/**
 * A single side of a <td>/<th>'s own border — either a per-side declaration
 * (`border-bottom:1px solid #hex`, used as a decorative divider line under a
 * row) or the plain `border`/`border-width`/`border-color`/`border-style`
 * shorthand applied uniformly to all four sides (`border:1px solid #999`,
 * the far more common case for an actual data table's cell borders). A
 * per-side declaration always wins over the all-sides one for that side,
 * same as real CSS. Independent of extractTableBorder, which only reads the
 * table's own border, not per-cell borders. `side` is one of
 * 'top'/'right'/'bottom'/'left'.
 */
export function extractCellBorder(styleAttr, side) {
  return extractSideBorder(styleAttr, side, true);
}

/**
 * A single side of a plain block element's (e.g. <div>) own border — reads
 * ONLY a genuinely side-specific declaration (`border-left:5px solid #hex`,
 * the common "accent stripe" pattern: a colored bar down one edge with no
 * border on the other three) with no fallback to the generic `border`
 * shorthand, since that uniform case is already covered by
 * extractTableBorder (reused for block elements — see dom-to-idm.js's
 * mapElement). `side` is one of 'top'/'right'/'bottom'/'left'.
 */
export function extractBlockSideBorder(styleAttr, side) {
  return extractSideBorder(styleAttr, side, false);
}

// Returns a plain 6-digit hex string (no '#'), or undefined if unrecognized.
// Exported for svg-shapes.js's fill/stroke resolution — an SVG presentation
// attribute (`fill="orange"`, `stroke="#2563eb"`) is the exact same color
// grammar as a CSS value, just carried on a bare attribute instead of a
// `style` declaration.
export function normalizeColor(value) {
  if (!value) return undefined;
  const v = value.trim().toLowerCase();

  const hexMatch = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(v);
  if (hexMatch) {
    const hex = hexMatch[1];
    return (hex.length === 3 ? hex.split('').map((c) => c + c).join('') : hex).toUpperCase();
  }

  const rgbMatch = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(v);
  if (rgbMatch) {
    return rgbMatch.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  return NAMED_COLORS[v]?.toUpperCase();
}

// Pure unit conversion (px/pt → pt), no keywords — shared by font-size and
// border-width parsing. Kept separate from normalizeFontSizePt's keyword
// table below because border-width has its own keywords ('thin'/'medium'/
// 'thick') that would otherwise collide with font-size's ('medium' means a
// very different size in each context).
function parseAbsoluteLengthPt(value) {
  if (!value) return undefined;
  const v = value.trim().toLowerCase();
  if (v === '0') return 0; // CSS allows a bare zero length with no unit
  const match = /^([\d.]+)(px|pt)$/.exec(v);
  if (!match) return undefined;
  const [, num, unit] = match;
  const n = Number(num);
  return unit === 'pt' ? n : n * 0.75; // 96 CSS px/in ÷ 72 pt/in
}

// The CSS absolute-size keywords are not relative to anything (unlike
// em/%/larger/smaller, which depend on an inherited base size this parser
// doesn't track) — the spec assigns each one a fixed size off a 16px medium,
// the same fixed px values every major browser actually renders.
const ABSOLUTE_FONT_SIZE_PX = {
  'xx-small': 9, 'x-small': 10, small: 13, medium: 16,
  large: 18, 'x-large': 24, 'xx-large': 32, 'xxx-large': 48,
};

// Handles absolute units (px/pt) and the absolute-size keywords above —
// relative units (em/%/larger/smaller) have no single correct conversion
// without knowing the inherited base size, so they're left unset rather
// than guessed.
function normalizeFontSizePt(value) {
  if (!value) return undefined;
  const v = value.trim().toLowerCase();
  if (v in ABSOLUTE_FONT_SIZE_PX) return ABSOLUTE_FONT_SIZE_PX[v] * 0.75;
  return parseAbsoluteLengthPt(v);
}
