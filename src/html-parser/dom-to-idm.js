/**
 * DOM → IDM mapper.
 *
 * Walks the DOM tree produced by the tree builder and converts each element to
 * its IDM equivalent. Noise (script, style, nav, event handlers) is dropped
 * silently. Metadata is extracted from <head>.
 */

import {
  code,
  codeSpan,
  doc,
  emphasis,
  figure,
  heading,
  image,
  link,
  list,
  listItem,
  paragraph,
  quote,
  section,
  strong,
  styledSpan,
  svgShapes,
  tableCell,
  tableRow,
  text,
} from '../idm/index.js';
import { extractRunStyle, extractAlign, extractVerticalAlign, extractDirection, extractWritingMode, extractRotationDeg, extractBackgroundColor, extractBackgroundGradient, extractTableBorder, extractBlockSideBorder, extractCellBorder, extractHeightPx, extractWidthPx, extractTableWidthPercent, extractMarginAuto, extractMaxWidthPx, extractMaxHeightPx, extractOverflowX, extractOverflowY, extractMarginPt, extractPaddingPt, extractDisplayMode, extractGridColumnCount, extractGapPt, extractLineHeight, isInlineBlock, hasBoxShadow, hasBorderRadius, isCircularBorderRadius, hasExplicitNoBorder, isHidden, preservesWhitespace, normalizeColor, parseInlineStyleAttr } from './inline-style.js';

/** Elements to drop entirely (no content propagated). */
const DROP_ELEMENTS = new Set([
  'script', 'style', 'noscript', 'nav', 'iframe',
  'template', 'head', 'meta', 'link', 'base', 'title',
]);

/**
 * Tags that make a <td>/<th>'s content block-level (HTML5 "flow content")
 * rather than plain phrasing content — most commonly a nested <table> in
 * table-based layouts (email templates, older CMS exports).
 */
const CELL_BLOCK_TAGS = new Set([
  'table', 'div', 'p', 'ul', 'ol', 'blockquote', 'pre',
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'section', 'article', 'figure', 'header', 'footer', 'aside', 'hr',
  // Flow-content wrapper tags (HTML5 permits a <p> inside any of these) that
  // previously weren't recognized as block-level at all — an unrecognized
  // tag falls through mapBlocks into the *inline* run path (see
  // mapInlineElementCore's default case), which has no concept of paragraph
  // breaks, so e.g. `<form><p>A</p><p>B</p></form>` silently merged into one
  // run of text ("AB") with no separator. Treating them exactly like
  // div/section above fixes that.
  'form', 'fieldset', 'address', 'details', 'summary', 'legend',
  // <canvas>'s own fallback content (shown only to browsers/renderers with no
  // canvas support) and <slot>'s own light-DOM/default content are both
  // ordinary flow content that can contain a <p> — the exact same merge risk
  // as form/fieldset above, just via a different real-world source (a canvas
  // chart with a text fallback, or raw pre-Shadow-DOM HTML using <slot>).
  // Neither tag's own defining behavior (script-driven drawing; Shadow DOM
  // content projection) has any document-export equivalent, so treating
  // whatever content they carry as plain block content is the correct
  // fallback, same as it is for form/fieldset.
  'canvas', 'slot',
]);

export function domToIdm(root) {
  const classRules = extractClassRules(root);
  if (classRules.length > 0) applyClassStyles(root, classRules);
  const metadata = extractMetadata(root);
  // Every <style> block's raw CSS text, concatenated — a font resolver (see
  // fonts/stylesheet-resolver.js) scans this for @font-face/@import rules
  // completely separately from applyClassStyles's own .class/tag selector
  // matching above; carried as one plain string on metadata rather than
  // re-walking the tree a second time downstream.
  const inlineStylesheetCss = collectStyleBlocksText(root);
  if (inlineStylesheetCss) metadata.inlineStylesheetCss = inlineStylesheetCss;
  const body = findBody(root) ?? root;
  // `body{font-family:Arial}` is a page-wide default, not a rule matched
  // per-element — this project's CSS resolution (applyClassStyles above) only
  // ever rewrites the *matching* element's own style attribute, it doesn't
  // implement real CSS inheritance from an ancestor to every descendant.
  // <body> itself is never mapped to an IDM node (only its children are), so
  // without reading it here explicitly, its own declared font never reaches
  // anything — every exporter's own hardcoded default typeface (DOCX/XLSX/
  // PPTX all default to Calibri) wins even when the source document clearly
  // declared a different one. Carried on `metadata.fontFamily` as a
  // document-wide fallback; any element with its own more specific
  // font-family (an explicit style or a matched class/tag rule) still wins,
  // same "explicit beats inferred" precedence used everywhere else.
  const bodyRunStyle = extractRunStyle(body?.attrs?.style);
  if (bodyRunStyle?.fontFamily) metadata.fontFamily = bodyRunStyle.fontFamily;
  // Same "document-wide fallback, an explicit per-element value still wins"
  // reasoning as fontFamily above, for the two other properties that
  // determine every text-block's own height (line count aside): a page
  // declaring `body{font:16px/1.6 ...}` (whether spelled out as `font-size`/
  // `line-height` longhands or the `font` shorthand — see inline-style.js's
  // own expandFontShorthand) expects ITS OWN base size/spacing to size
  // every plain paragraph/list/table-cell/etc, not this project's own
  // hardcoded default (DOCX/PDF's Normal-style-equivalent 11pt/1.2). Without
  // this, a document built around a deliberately larger or more
  // tightly-spaced base font renders every unstyled block at the wrong
  // size regardless of what the page itself declared.
  if (bodyRunStyle?.fontSizePt) metadata.fontSizePt = bodyRunStyle.fontSizePt;
  const bodyLineHeight = extractLineHeight(body?.attrs?.style);
  if (bodyLineHeight) metadata.lineHeight = bodyLineHeight;
  // `body{max-width:1000px;margin:auto}` is the standard "centered page-like
  // container" pattern — it's the actual width a browser renders the page
  // at, regardless of viewport/window size. A paged exporter (PDF) has no
  // equivalent of an unbounded scrolling viewport, so without reading this,
  // it always falls back to a fixed Letter/A4 width that has no relationship
  // to what the source page actually looks like — carried on
  // `metadata.contentWidthPx` so a paged exporter CAN size its own page to
  // match instead, though isn't required to (a non-paged exporter, or a
  // paged one with an explicit user-chosen page size, ignores it entirely —
  // same "explicit beats inferred" precedence as fontFamily/fontSizePt
  // above). If EVERY top-level element child of body declares the same own
  // left+right padding — the common `body > .page{padding:40px}` pattern,
  // including its equally common multi-section form (one `.page` div per
  // eventual PDF page, e.g. via a `.pagebreak` class on all but the first —
  // several siblings, all still sharing one `.page` rule) — that shared
  // padding further narrows the effective content width, the same way it
  // would inset any other padded box. Deliberately requires EVERY child to
  // agree on the exact same value, not just the first one: children that
  // don't all play the same "page" role (mismatched or absent padding on
  // even one of them) give no reliable single answer, and guessing which
  // one is "the real" page would be exactly as likely wrong as right.
  const bodyElementChildren = (body.children ?? []).filter((n) => n.type === 'element' && !DROP_ELEMENTS.has(n.name));
  // A bare `<body>` declaring its own max-width directly is actually the LESS
  // common real-world form of this pattern — far more often the whole page is
  // wrapped in a single container `<div style="width:1200px;max-width:100%;
  // margin:0 auto">` instead (Bootstrap's `.container`, hand-rolled invoice/
  // email templates, ...), with `<body>` itself carrying no sizing at all.
  // When body has exactly one such wrapper child and body itself declares no
  // max-width, that wrapper's own width is used instead — preferring its
  // `max-width` if set, else its `width` (the actual design size; an
  // accompanying `max-width:100%` alongside `width` is just a responsiveness
  // guard for narrow viewports, not the real target width). The padding-based
  // narrowing below then also switches to checking THAT wrapper's own
  // children (the real per-"page" elements) rather than body's.
  let widthOwner = body;
  let widthOwnerPx = extractMaxWidthPx(body?.attrs?.style);
  if (widthOwnerPx === undefined && bodyElementChildren.length === 1) {
    const wrapper = bodyElementChildren[0];
    const wrapperWidthPx = extractMaxWidthPx(wrapper.attrs?.style) ?? extractWidthPx(wrapper.attrs?.style);
    if (wrapperWidthPx) {
      widthOwner = wrapper;
      widthOwnerPx = wrapperWidthPx;
    }
  }
  if (widthOwnerPx) {
    let contentWidthPx = widthOwnerPx;
    const widthOwnerChildren = widthOwner === body ? bodyElementChildren
      : (widthOwner.children ?? []).filter((n) => n.type === 'element' && !DROP_ELEMENTS.has(n.name));
    if (widthOwnerChildren.length >= 1) {
      const paddings = widthOwnerChildren.map((el) => extractPaddingPt(el.attrs?.style)).filter(Boolean);
      if (paddings.length === widthOwnerChildren.length) {
        const widths = paddings.map((p) => (p.left ?? 0) + (p.right ?? 0));
        if (widths.every((w) => w === widths[0]) && widths[0] > 0) {
          contentWidthPx -= widths[0] / 0.75; // pt -> px
          // The SAME shared padding this project already uses as the page's
          // own left/right margin — a browser has no separate "print
          // margin" concept layered on top of `.page`'s own edges, so
          // stacking this project's usual 1in default margin on top of an
          // already-narrowed content width would add whitespace the source
          // page never had. Halved onto each side, matching the symmetric
          // shorthand `padding:Npx` this pattern is written with in
          // practice (an asymmetric page padding is not this pattern).
          metadata.pageMarginPt = widths[0] / 2;
        }
      }
    }
    if (contentWidthPx > 0) metadata.contentWidthPx = contentWidthPx;
  }
  const children = mapBlocks(body.children);
  return doc(children, metadata);
}

/* -------------------------------------------------------------------------- */
/* <style> block → class/tag attribute resolution                             */
/* -------------------------------------------------------------------------- */
// A very common real-world pattern this project didn't handle at all: CSS
// declared once in a <style> block and applied via class="..." or a bare tag
// name, rather than repeated inline on every element. Resolved forms:
// `.name { ... }`, `tagname { ... }`, a comma list of such, and a two-part
// descendant OR child combinator of the same atoms (`.box h1 { ... }`,
// `.box .title { ... }`, `.flex > div { ... }` — the "give this specific
// override to whatever's inside the box" pattern that near-always
// accompanies a `.box`-level background/color rule in real templates; `>` is
// treated identically to a plain descendant space, matching any depth rather
// than strictly a direct child). Tag+class compounds (`div.box`), 3+-part
// chains, sibling combinators (`+`, `~`), and pseudo-classes (:nth-child,
// etc.) need real CSS cascade/specificity resolution well beyond "carry
// forward an explicit declaration," so those rules are left unmatched, same
// as if absent. This runs as a pre-pass that rewrites each matched element's
// own `style` attribute in place — every existing style-extraction function
// downstream keeps reading `el.attrs.style` exactly as before, unchanged.

const CLASS_SELECTOR = /^\.([a-zA-Z0-9_-]+)$/;
const TAG_SELECTOR = /^[a-zA-Z][a-zA-Z0-9-]*$/;

function extractClassRules(root) {
  const rules = []; // { kind: 'class' | 'tag', key, declarations } | { kind: 'descendant', ancestor, descendant, declarations }, in stylesheet order
  walkForStyleBlocks(root, rules);
  return rules;
}

function walkForStyleBlocks(node, rules) {
  if (node.type !== 'element') return;
  if (node.name === 'style') {
    parseClassRules(textOf(node), rules);
    return; // <style>'s raw-text content has no element children to recurse into
  }
  for (const child of node.children) walkForStyleBlocks(child, rules);
}

// A single class-or-tag atom, e.g. one side of a descendant combinator —
// returns { kind, key } or undefined if it's neither (an id/attribute/
// pseudo-class/compound selector, which this resolver doesn't model).
function matchSelectorAtom(token) {
  const classMatch = CLASS_SELECTOR.exec(token);
  if (classMatch) return { kind: 'class', key: classMatch[1] };
  if (TAG_SELECTOR.test(token)) return { kind: 'tag', key: token.toLowerCase() };
  return undefined;
}

function parseClassRules(css, rules) {
  const ruleRegex = /([^{}]+)\{([^{}]*)\}/g;
  let match;
  while ((match = ruleRegex.exec(css))) {
    const [, selectorList, declarations] = match;
    for (const selector of selectorList.split(',')) {
      const trimmed = selector.trim();
      const single = matchSelectorAtom(trimmed);
      if (single) {
        rules.push({ ...single, declarations });
        continue;
      }
      // Descendant combinator: exactly two atoms (`.box h1`, `.box .title`,
      // `.box > div`) — anything else (3+ parts, `+`/`~`, a compound like
      // `div.box`) doesn't split into two plain atoms and is left unmatched
      // below. A child combinator (`.flex > div`, with or without spaces
      // around the `>`) is normalized to a plain descendant combinator
      // first: matching any depth rather than strictly a direct child is
      // already the exact approximation elementMatchesAtom/applyClassStyles
      // makes for a real descendant selector (see their own note on
      // "matches at any depth"), so treating `>` the same way costs nothing
      // extra — it just stops a very common real-world pattern ("style my
      // own direct children", e.g. `.flex > div{background:...}` sitting
      // right next to `.grid div{background:...}` in the same stylesheet)
      // from silently resolving one and dropping the other.
      const parts = trimmed.replace(/>/g, ' ').split(/\s+/).filter(Boolean);
      if (parts.length === 2) {
        const ancestor = matchSelectorAtom(parts[0]);
        const descendant = matchSelectorAtom(parts[1]);
        if (ancestor && descendant) rules.push({ kind: 'descendant', ancestor, descendant, declarations });
      }
    }
  }
}

// Same whole-tree <style> walk as walkForStyleBlocks above, kept separate
// rather than merged into it — that function's job is producing matchable
// selector rules (a very different, class/tag-selector-specific shape) for
// applyClassStyles; this one just wants every block's raw text, verbatim,
// for a completely different consumer (@font-face/@import scanning).
function collectStyleBlocksText(node, blocks = []) {
  if (node.type !== 'element') return blocks.join('\n');
  if (node.name === 'style') {
    blocks.push(textOf(node));
    return blocks.join('\n');
  }
  for (const child of node.children) collectStyleBlocksText(child, blocks);
  return blocks.join('\n');
}

function elementMatchesAtom(el, atom) {
  if (atom.kind === 'tag') return el.name === atom.key;
  const elementClasses = (el.attrs['class'] ?? '').split(/\s+/).filter(Boolean);
  return elementClasses.includes(atom.key);
}

// Rewrites every element's own style attribute to `matchedRuleDeclarations;
// originalInlineStyle` — inline always comes last, so it wins over class/tag
// declarations for the same property (parseInlineStyleAttr already applies
// "last declaration for a property wins"), matching real CSS specificity.
// Matched rules are concatenated in *stylesheet definition order* (not the
// order classes are listed in the class="..." attribute — real CSS
// specificity is source order, not attribute order), so a later-defined
// rule correctly overrides an earlier one for a shared property, regardless
// of whether it targeted the element by class, tag name, or descendant combinator.
//
// `activeAncestors` is the set of rule indexes (into `rules`) whose ancestor
// half is already satisfied by some strict ancestor of `node` — a descendant
// combinator matches at any depth, not just a direct child, so once an
// ancestor satisfies it, every descendant carries that candidacy forward
// (`childActive` below), not just node's immediate children.
function applyClassStyles(node, rules, activeAncestors = new Set()) {
  if (node.type !== 'element') return;

  const childActive = new Set(activeAncestors);
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (rule.kind === 'descendant' && elementMatchesAtom(node, rule.ancestor)) childActive.add(i);
  }

  if (rules.length > 0) {
    const matched = [];
    for (let i = 0; i < rules.length; i++) {
      const rule = rules[i];
      if (rule.kind === 'descendant') {
        if (activeAncestors.has(i) && elementMatchesAtom(node, rule.descendant)) matched.push(rule.declarations);
      } else if (elementMatchesAtom(node, rule)) {
        matched.push(rule.declarations);
      }
    }
    if (matched.length) {
      const ruleStyle = matched.join(';');
      node.attrs.style = node.attrs?.style ? `${ruleStyle};${node.attrs.style}` : ruleStyle;
    }
  }

  for (const child of node.children) applyClassStyles(child, rules, childActive);
}

/* -------------------------------------------------------------------------- */
/* Metadata extraction                                                         */
/* -------------------------------------------------------------------------- */

function extractMetadata(root) {
  const metadata = {};

  // <html lang="en">
  const htmlEl = findChild(root, 'html');
  if (htmlEl?.attrs['lang']) metadata.language = htmlEl.attrs['lang'];

  // <head> is a child of <html>, not of #root
  const head = htmlEl ? findChild(htmlEl, 'head') : findChild(root, 'head');

  if (!head) return metadata;

  for (const child of head.children) {
    if (child.type !== 'element') continue;
    if (child.name === 'title') {
      metadata.title = textOf(child).trim() || undefined;
    }
    if (child.name === 'meta') {
      const name = (child.attrs['name'] ?? '').toLowerCase();
      const content = child.attrs['content'] ?? '';
      if (name === 'author') metadata.author = content;
    }
    // A narrow, deliberate exception to <link> otherwise being dropped
    // entirely (see DROP_ELEMENTS above) — mirrors the `title`/`meta`
    // carve-outs just above: almost every other `<link rel="...">` value
    // (icon/manifest/preload/canonical/...) has no bearing on document
    // export, but a stylesheet CAN declare @font-face rules a font resolver
    // needs (see fonts/stylesheet-resolver.js) — a plain URL string is all
    // that's carried here; fetching/parsing it happens entirely downstream.
    if (child.name === 'link' && (child.attrs['rel'] ?? '').toLowerCase() === 'stylesheet' && child.attrs['href']) {
      metadata.stylesheetLinks = [...(metadata.stylesheetLinks ?? []), child.attrs['href']];
    }
  }
  return metadata;
}

function findBody(root) {
  // Try <html><body>
  const html = findChild(root, 'html');
  if (html) {
    const body = findChild(html, 'body');
    if (body) return body;
    return html;
  }
  // Document fragment without <html>/<body>
  return undefined;
}

function findChild(el, name) {
  for (const child of el.children) {
    if (child.type === 'element' && child.name === name) return child;
  }
  return undefined;
}

/* -------------------------------------------------------------------------- */
/* Block mapping                                                                */
/* -------------------------------------------------------------------------- */

// Tags whose own mapping already produces a complete, standalone block (a
// heading, list, table, section, image, ...) — a superset of CELL_BLOCK_TAGS
// (which only needs the table-cell subset) since a top-level image/svg/dl
// is also always its own block here, unlike inside a <td> where it can sit
// inline with text. Anything else reaching mapBlocks — a bare text node, or
// an inline-level tag (<strong>/<em>/<a>/<span>/<code>/<mark>/<u>/<s>/<br>/
// <sub>/<sup>/...) — is phrasing content with no block of its own.
// <button> is included even though CSS treats it as inline-level: unlike a
// real inline element, a real-world <button> is essentially never embedded
// mid-sentence — it's a standalone CTA — so treating it as its own block
// lets mapElement's background/padding/border extraction actually apply to
// it (giving it a real padded, colored button box) instead of it being
// merged as plain phrasing content into whatever paragraph surrounds it,
// which only picks up run-level color/background/bold — never padding.
const TOP_LEVEL_BLOCK_TAGS = new Set([...CELL_BLOCK_TAGS, 'main', 'dl', 'dt', 'dd', 'img', 'svg', 'button']);

// Consecutive runs of phrasing content directly inside a block container
// (e.g. `<div class="note"><strong>Info:</strong> This is Word-friendly.</div>`
// — no wrapping <p>, very common in real-world CMS/hand-written HTML) collapse
// into ONE shared paragraph, the same way a browser treats a block element's
// directly-nested text/inline children as a single anonymous block box rather
// than a new line per child. Splitting each sibling into its own paragraph
// (the previous behavior) broke apart sentences like the one above into
// separate lines with no visual relationship to each other.
function mapBlocks(nodes) {
  const out = [];
  let inlineRun = [];

  const flushInlineRun = () => {
    if (inlineRun.length === 0) return;
    // A run consisting only of insignificant whitespace (e.g. the newline/
    // indentation between two sibling block elements in the source HTML)
    // produces no paragraph, matching the old behavior of silently dropping
    // an empty/whitespace-only bare text node.
    const meaningful = inlineRun.some((n) => n.type !== 'text' || n.value.trim() !== '');
    if (meaningful) {
      const content = mapInline(inlineRun);
      if (content.length) out.push(paragraph(content));
    }
    inlineRun = [];
  };

  for (const node of nodes) {
    if (node.type === 'element' && isHidden(node.attrs?.style)) continue;
    const isBlockElement = node.type === 'element' && (TOP_LEVEL_BLOCK_TAGS.has(node.name) || DROP_ELEMENTS.has(node.name));
    if (!isBlockElement) {
      inlineRun.push(node);
      continue;
    }
    flushInlineRun();
    if (DROP_ELEMENTS.has(node.name)) continue;
    out.push(...mapElement(node));
  }
  flushInlineRun();
  return out;
}

// A `display:flex`/`display:grid` container's children each become one
// positioned cell in PPTX (see pptx/slide.js's classifyContentUnits/
// renderLayoutBlock) instead of stacking as plain vertical text-flow the way
// mapBlocks' generic recursion would — the common "row of colored boxes" /
// "card grid" pattern. `layoutGroup` is a single shared, info-carrying
// token (`{ mode, columns, gapPt }` — object identity doubles as the
// grouping key, same trick dom-to-idm.js's mapElement uses for cardGroupId)
// stamped on every node each child contributes, alongside `layoutItemIndex`
// so multiple nodes from one child (e.g. a grid cell that's itself a
// multi-block card) stay together as one cell rather than becoming several.
// Every exporter besides PPTX simply never reads these two fields and keeps
// seeing the same flattened paragraph/heading sequence as before.
//
// Only direct block-level children become cells — a `display:flex` div
// wrapping a bare text node with no child elements at all (e.g. a circular
// badge whose own CSS uses `display:flex` purely to center its "72%" text,
// not to lay out multiple item children — a genuinely common pattern) has
// no block-level children to treat as cells. Falling back to plain
// mapBlocks recursion in that case is essential, not just a nicer default:
// the item-counting loop below only ever visits block-level children, so
// with no fallback a bare text child would silently vanish from the output
// instead of degrading to ordinary text-flow.
function mapLayoutBlocks(el, mode) {
  const layoutGroup = {
    mode,
    columns: mode === 'grid' ? extractGridColumnCount(el.attrs?.style) : undefined,
    gapPt: extractGapPt(el.attrs?.style),
  };
  const out = [];
  let itemIndex = 0;
  for (const child of el.children) {
    if (child.type !== 'element' || isHidden(child.attrs?.style)) continue;
    if (!TOP_LEVEL_BLOCK_TAGS.has(child.name) && !DROP_ELEMENTS.has(child.name)) continue;
    if (DROP_ELEMENTS.has(child.name)) continue;
    const childNodes = mapElement(child);
    if (childNodes.length === 0) continue;
    for (const node of childNodes) {
      if ((node.type === 'paragraph' || node.type === 'heading') && node.layoutGroup === undefined) {
        node.layoutGroup = layoutGroup;
        node.layoutItemIndex = itemIndex;
      }
    }
    out.push(...childNodes);
    itemIndex++;
  }
  return itemIndex > 0 ? out : mapBlocks(el.children);
}

// Applies paragraph-level style (text-align, border, background-color,
// vertical margin, padding, bookmark id) to whichever paragraph/heading
// nodes an element produced — the most common real-world case being a <div
// class="box"> (border+background) or a <div style="margin:10px 0">
// wrapping a single line of text. A <div>/<section> wrapping several blocks
// with its own style would need each child styled individually to be fully
// correct — out of scope for now; the common case (the element itself
// becoming exactly the styled paragraph/heading) is covered. A gradient
// background (`background: linear-gradient(...)`) has no flat-fill OOXML
// equivalent — extractBackgroundColor falls back to the gradient's first
// explicit color stop as an approximation. A CSS page-break-before/
// break-before (e.g. a bare `<div class="pagebreak">`) inserts a real page
// break ahead of whatever the element produced, same as `<hr
// class="page-break">` already does.
//
// `nodes` here is the *flattened* result of every descendant block this
// element wraps — for a container that nests further styled containers
// (e.g. a "page" div wrapping a "hero" box), a paragraph/heading several
// levels down may already carry a value a nested wrapper set for itself
// (mapElement runs bottom-up, so the inner wrapper's assignment always
// happens first). Only filling in a field when it's still unset means the
// most deeply-nested (most specific) wrapper's styling wins, matching CSS's
// "background doesn't cascade" behavior — without this guard, an outer
// container's own background/padding/margin overwrites every already-styled
// nested box's, which is actively wrong rather than just incomplete.
//
// `color`, unlike those, genuinely *is* an inherited CSS property: a box
// that sets its own text color (`.hero{background:#1f4e79;color:#fff}`, the
// classic "colored panel" pattern) expects every plain child paragraph to
// pick it up even when nothing overrides it explicitly — the common case
// being a heading with its own more-specific color rule (`.hero h1{color:
// #fff}`) sitting next to a plain <p> that only relies on the container's
// color. It's carried by wrapping the child's *inline content* (not a flat
// field like the others) in an extra StyledSpanNode layer, so it composes
// safely with whatever coloring that content already has: at render time an
// inner, already-baked styledSpan's own color always wins over one added by
// an outer wrapper (see document.js's renderRun — later-applied/more-nested
// rProps override earlier ones), which is exactly the same
// most-specific-wins precedence used for background/padding/margin above.
//
// Only computed for the pure-recursion container tags (div/header/footer/
// aside/dl/form/fieldset/address/details/summary/legend/canvas/slot — see
// mapElementCore's "just recurse" cases): every other tag that can reach
// here (h1-h6, p, dt, dd, the inline-content fallback) already runs its own
// el.attrs.style through mapInlineWithStyle/mapInline while building `nodes`,
// so re-extracting color from that same el here would double-wrap it in a
// second, redundant StyledSpanNode.
const COLOR_INHERITING_CONTAINERS = new Set([
  'div', 'header', 'footer', 'aside', 'dl',
  'form', 'fieldset', 'address', 'details', 'summary', 'legend',
  'canvas', 'slot',
]);

// Every IDM node type a wrapping element's own border/background/padding/etc.
// can bleed onto — the single list to extend whenever a new node type turns
// out to be missing this (this project's own repeated bug: `<blockquote>`
// silently dropped a wrapping `<div>`'s styling because 'quote' wasn't
// checked here, and the same turned out to be true of 'list'/'table'/'code'/
// 'figure'/'section' too — each discovered as its own separate report before
// this was generalized). Adding a type here is now the *entire* fix for "this
// node's own CSS box styling gets silently dropped by mapElement" — no new
// hand-written branch needed, which is exactly the class of bug this set
// exists to stop from recurring one node type at a time.
const BOX_STYLE_NODE_TYPES = new Set(['paragraph', 'heading', 'quote', 'section', 'list', 'table', 'code', 'figure']);

// Of BOX_STYLE_NODE_TYPES above, only these two have genuinely *inline*
// `children` (see normalizer's normalizeBlock/normalizeInline — every other
// type in that set has block children instead, e.g. a quote's children are
// paragraph nodes). `inheritedColor`'s own handling below wraps `children` in
// an inline StyledSpanNode, which is only safe for these two — applying it to
// a quote/list/table/etc. would corrupt their block children into being an
// inline node's children instead.
const INLINE_CHILDREN_NODE_TYPES = new Set(['paragraph', 'heading']);

function mapElement(el) {
  const nodes = mapElementCore(el);
  const align = extractAlign(el.attrs?.style);
  const border = extractTableBorder(el.attrs?.style); // generic border reader, not table-specific
  // A per-side border (e.g. `.note{border-left:5px solid #hex}`, the common
  // "accent stripe" pattern) — independent of the uniform `border` above,
  // which only covers a border applied equally to all four sides. See
  // extractBlockSideBorder for why this doesn't also fall back to the
  // generic `border` shorthand (that's what `border` above already handles).
  const borderTop = extractBlockSideBorder(el.attrs?.style, 'top');
  const borderRight = extractBlockSideBorder(el.attrs?.style, 'right');
  const borderBottom = extractBlockSideBorder(el.attrs?.style, 'bottom');
  const borderLeft = extractBlockSideBorder(el.attrs?.style, 'left');
  const backgroundColor = extractBackgroundColor(el.attrs?.style);
  // The full multi-stop gradient, kept alongside backgroundColor's
  // single-color approximation — see extractBackgroundGradient for why only
  // PPTX (via a real `a:gradFill`) can use it; every other exporter just
  // keeps reading backgroundColor and never sees this field.
  const backgroundGradient = extractBackgroundGradient(el.attrs?.style);
  const margin = extractMarginPt(el.attrs?.style);
  const padding = extractPaddingPt(el.attrs?.style);
  const anchorId = el.attrs?.id;
  const inheritedColor = COLOR_INHERITING_CONTAINERS.has(el.name) ? extractRunStyle(el.attrs?.style)?.color : undefined;
  // `display:inline-block`/an explicit fixed width are the real CSS signals
  // that a short bordered/filled box should shrink to its content (or to
  // that width) instead of stretching to fill its container — see
  // pptx/shapes.js's isStyledBoxNode, which used to guess this purely from
  // how short the text was, wrongly treating a full-width gradient button
  // (no inline-block, no width — just a plain block-level <div>) the same
  // as a real shrink-wrapped badge/pill.
  //
  // An explicit `width` paired with auto left+right margins (`margin:0
  // auto`, or the equally common `margin-left:auto;margin-right:auto`) is a
  // DIFFERENT, well-known idiom — the centered fluid "page container"
  // pattern (`.container{width:1200px;max-width:100%;margin:0 auto}`,
  // Bootstrap's `.container` and countless hand-rolled invoice/email
  // templates), not a shrink-wrapped badge: a real badge/button is never
  // itself horizontally auto-centered inside its own parent this way, it's
  // sized to its own width and left where it sits. Without this exclusion, a
  // page-wide container's own `width` bled `shrinkToFit` (see the bleed loop
  // below) onto EVERY paragraph/heading anywhere in the document — including
  // ones with their own genuine backgroundColor/border further down (a
  // right-aligned address block, a centered full-width banner) — silently
  // shrinking each one to its own tight content box instead of the page's
  // full width, and (for a right/center-aligned one) risking positioning it
  // far off the page entirely (see pdf/blocks.js's renderStyledRunBlock).
  const widthPx = extractWidthPx(el.attrs?.style);
  const heightPx = extractHeightPx(el.attrs?.style);
  const isCenteredContainer = widthPx != null && extractMarginAuto(el.attrs?.style) === 'both';
  const shrinkToFit = isInlineBlock(el.attrs?.style) || (widthPx != null && !isCenteredContainer) || undefined;
  // DrawingML can't reproduce a CSS shadow's own offset/blur/color, only its
  // presence — see hasBoxShadow / pptx/shapes.js's shadowEffectXml.
  const shadow = hasBoxShadow(el.attrs?.style) || undefined;
  // A real CSS `border-radius` — OOXML paragraph/table borders can't be
  // rounded, so this is only ever a presence signal (see hasBorderRadius)
  // used to pick a rounded-corner DrawingML shape (`roundRect`) over a
  // square one when a box is promoted to a real shape. Gated the same way
  // as `shadow` below (not a plain "if unset" bleed like border/background):
  // a page-level wrapper's own `border-radius` (e.g. `.slide{border-radius:
  // 15px}`) must not round every already-independently-styled child inside
  // it, only elements that are genuinely rounded themselves.
  const rounded = hasBorderRadius(el.attrs?.style) || undefined;
  // `border-radius:50%` (or more) — the standard CSS circle/oval idiom — is
  // a stronger, more specific signal than a plain rounded corner: it means
  // the box should become a real `ellipse`, not just a `roundRect` with
  // slightly softened corners. Gated the same way as `rounded`/`shadow`
  // above, for the same reason (a page-level wrapper's own border-radius
  // must not turn every child inside it into a circle).
  const circular = isCircularBorderRadius(el.attrs?.style) || undefined;
  // A container wrapping several block children (a "card": a heading plus a
  // couple of paragraphs, all inside one bordered/filled <div>) has no
  // single paragraph/heading to stamp the border/fill onto the way the
  // one-child case below does — every exporter that only understands
  // per-paragraph border/shading (DOCX, XLSX) still gets a reasonable
  // approximation from that per-node stamping, but PPTX needs to know these
  // particular nodes came from one shared box so it can draw one shape
  // behind all of them instead of nothing at all (see pptx/slide.js's
  // classifyContentUnits). `cardGroupId` is just a unique per-container token —
  // object identity, not a value — shared by every node this container
  // produced, guarded the same "only if unset" way as the fields below so a
  // more deeply nested card's own grouping always wins.
  //
  // Requires an actual `border` — deliberately narrower than backgroundColor/
  // backgroundGradient alone: a border is a strong, specific "this is one
  // bounded panel" signal, but a plain background is exactly as likely to be
  // a big page-level wrapper's own backdrop (`.slide{background:white;
  // box-shadow:...}` wrapping an entire slide's worth of title/badge/
  // paragraph/list/table content) as it is a real card. Multi-node
  // promotion on backgroundColor alone turned every otherwise-plain
  // paragraph inside such a wrapper into a phantom bordered/shadowed box it
  // was never meant to have — a background-only container still bleeds its
  // color onto unstyled children as a plain run highlight (see the
  // backgroundColor assignment below), just without becoming its own shape.
  //
  // `backgroundColor && inheritedColor` together is trusted as an equally
  // strong second signal, alongside border: a deliberate colored panel (a
  // `.hero{background:#1f4e79;color:#fff}` banner — a genuinely common
  // pattern, a heading + a short subtitle both sitting on a colored
  // background) reliably pairs its own background with an explicit
  // contrasting text color, precisely because plain dark-on-light text would
  // be unreadable on it. A big neutral page-level wrapper
  // (`.page{background:#fff}`) essentially never does this — a white/plain
  // backdrop needs no color override at all, so `color` stays unset. Without
  // this, `.hero`'s heading+paragraph had neither signal (no border, no
  // shadow) and fell all the way through to PPTX's plain run-highlight
  // fallback: two separate tight-highlighted lines with a gap between them
  // instead of one shared colored box.
  const trustedPanel = Boolean(border) || Boolean(backgroundColor && inheritedColor);
  const cardGroupId = nodes.length > 1 && trustedPanel ? {} : undefined;
  // Whether shadow/rounded/circular can be trusted as genuinely describing
  // *this* wrapper's own box rather than a big multi-child page-level
  // wrapper's backdrop bleeding onto one coincidentally-plain child — see the
  // shadow/rounded/circular assignments below for the full reasoning. Mirrors
  // cardGroupId's own "one bounded panel" criterion (now including the
  // backgroundColor+inheritedColor signal too) exactly.
  const shadowRoundedTrusted = nodes.length === 1 || trustedPanel;
  // A wrapping container's own margin is external to its box — it belongs
  // above the FIRST child and below the LAST child, never repeated on every
  // child in between. Unlike padding/background/border (which this project
  // deliberately treats as "themed defaults an unclaimed child inherits",
  // see the paddingAlreadyOwned notes below and its own test coverage),
  // margin has no such "inherits the container's look" justification — a
  // `.page{margin:24px 0}` wrapping ten blocks means one 24px gap above the
  // whole page and one below it, not ten repeated 24px gaps between every
  // sibling. Computed once per element using the EXACT SAME eligibility the
  // assignment loop below applies — BOX_STYLE_NODE_TYPES *and* not already
  // owned by some inner wrapper (replicating paddingAlreadyOwned's own
  // fields here, since that flag is only computed per-node inside the loop
  // below). Skipping the ownership half would pick, say, a hero banner's own
  // `<h1>` as "the first child" whenever that hero happens to be a
  // container's very first block (extremely common — a colored title banner
  // at the top of a "page"/"card"), then the real assignment below refuses
  // it anyway (`!paddingAlreadyOwned`), silently dropping the container's
  // entire margin-top instead of landing on the next actual candidate.
  const marginEligibleNodes = margin
    ? nodes.filter((n) => BOX_STYLE_NODE_TYPES.has(n.type) && n.border === undefined && n.backgroundColor === undefined &&
        n.backgroundGradient === undefined && n.cardGroup === undefined && n.layoutGroup === undefined)
    : [];
  const firstMarginNode = marginEligibleNodes[0];
  const lastMarginNode = marginEligibleNodes[marginEligibleNodes.length - 1];
  if (align || border || borderTop || borderRight || borderBottom || borderLeft || backgroundColor || backgroundGradient || margin || padding || anchorId || inheritedColor || shrinkToFit || shadow || rounded || circular) {
    for (const node of nodes) {
      if (!BOX_STYLE_NODE_TYPES.has(node.type)) continue;
      // Whether some *inner*, more specific wrapper already gave this node
      // its own border/fill/grouping before this container's turn — e.g. a
      // `<div class="gradient">...</div>` that already fully styled
      // itself, sitting as one of several children inside an outer
      // `<div class="slide">` that also happens to have its own
      // background. That already-self-contained node must NOT also be
      // swept into the outer container's cardGroup just because its own
      // cardGroup field happens to still be unset (a single-child
      // container never creates one for itself) — doing so would silently
      // reclassify an already-correct standalone badge as a one-node
      // "card", changing how PPTX renders it for no reason.
      const styledByInnerWrapper = node.border !== undefined || node.backgroundColor !== undefined ||
        node.backgroundGradient !== undefined || node.cardGroup !== undefined || node.layoutGroup !== undefined ||
        node.shrinkToFit !== undefined;
      // Deliberately narrower than styledByInnerWrapper above: excludes
      // shrinkToFit, which — unlike border/backgroundColor/cardGroup/
      // layoutGroup, all of which are only ever set by a bleed-down loop
      // like this one, never by mapElementCore itself — CAN be set
      // directly by mapElementCore for the current element in the same
      // call (the <button> case). Using the wider flag here would treat a
      // button's own shrinkToFit (set moments ago, same element, same
      // call) as if some deeper wrapper had already claimed it, blocking
      // the button's own padding/margin from ever attaching to its own
      // node. border/backgroundColor/etc. don't have that same-call race,
      // so they still correctly detect a genuinely deeper, already
      // self-contained wrapper (e.g. a circular badge that already set its
      // own backgroundColor before an outer `.slide{padding:...}` gets its
      // turn) without the button false positive.
      const paddingAlreadyOwned = node.border !== undefined || node.backgroundColor !== undefined ||
        node.backgroundGradient !== undefined || node.cardGroup !== undefined || node.layoutGroup !== undefined;
      if (align && node.align === undefined) node.align = align;
      if (border && node.border === undefined) node.border = border;
      if (borderTop && node.borderTop === undefined) node.borderTop = borderTop;
      if (borderRight && node.borderRight === undefined) node.borderRight = borderRight;
      if (borderBottom && node.borderBottom === undefined) node.borderBottom = borderBottom;
      if (borderLeft && node.borderLeft === undefined) node.borderLeft = borderLeft;
      if (backgroundColor && node.backgroundColor === undefined) node.backgroundColor = backgroundColor;
      if (backgroundGradient && node.backgroundGradient === undefined) node.backgroundGradient = backgroundGradient;
      // Gated on !paddingAlreadyOwned (see its own note above): padding/
      // margin are box-model properties of the specific element that
      // declared them, never inherited by descendants in real CSS —
      // bleeding an outer container's own padding (e.g. `.slide{padding:
      // 40px}`) down through several nesting levels onto an already
      // self-contained shrink-to-fit badge/circle would inflate that
      // badge's tiny fixed-size shape with padding many times larger than
      // its own diameter, leaving almost no room for its label text inside
      // the shape (observed as a circular badge's text force-wrapping to
      // one character per line in Word). A plain `<div style="padding:
      // 20px"><p>Text</p></div>` — the case this bleed exists for — still
      // works: the inner <p> has no border/background/shrinkToFit of its
      // own, so paddingAlreadyOwned is false for it.
      if (margin && !paddingAlreadyOwned && node.margin === undefined) {
        // A single eligible child (the common case this bleed exists for —
        // see the padding note just above) still gets both sides, exactly as
        // before; only a multi-child wrapper splits before/after across its
        // first/last child instead of repeating both on every one of them.
        const beforePt = node === firstMarginNode ? margin.beforePt : undefined;
        const afterPt = node === lastMarginNode ? margin.afterPt : undefined;
        if (beforePt !== undefined || afterPt !== undefined) node.margin = { beforePt, afterPt };
      }
      // ALSO gated on shadowRoundedTrusted (the same "genuinely one bounded
      // box" signal shadow/rounded/circular already require below): a
      // multi-child, untrusted wrapper's own padding (e.g. `.page{padding:
      // 40px}` wrapping ten blocks) describes room *inside that one
      // container's own box*, not room inside each individual child's own
      // box. Without this, a bled `backgroundColor` (still allowed through
      // unconditionally above, as a deliberate "plain highlight" — see its
      // own note) makes nodeBoxStyle's `hasBox` true for every such child
      // too (PDF's blocks.js/DOCX's document.js/PPTX's shapes.js all read
      // `hasBox` as "draw a real padded box"), so each child ended up with
      // its own copy of the *entire* container padding stacked on top of the
      // next — the single biggest source of runaway vertical whitespace in a
      // real "page" or "card" wrapper with several children. A single child
      // (nodes.length === 1) or a genuinely trusted panel (an explicit
      // border, or backgroundColor+color together) still gets it, unchanged.
      if (padding && shadowRoundedTrusted && !paddingAlreadyOwned && node.padding === undefined) node.padding = padding;
      if (anchorId && node.anchorId === undefined) node.anchorId = anchorId;
      if (cardGroupId && !styledByInnerWrapper) node.cardGroup = cardGroupId;
      if (shrinkToFit && node.shrinkToFit === undefined) node.shrinkToFit = shrinkToFit;
      if (widthPx != null && node.widthPx === undefined) node.widthPx = widthPx;
      if (heightPx != null && node.heightPx === undefined) node.heightPx = heightPx;
      // Gated on !styledByInnerWrapper (unlike border/backgroundColor's own
      // plain "if unset" bleed, which is a deliberate, accepted "sits on
      // this backdrop" approximation) — box-shadow isn't a subtle backdrop
      // effect, it's a strong, visible "this specific element floats above
      // the page" signal, and bleeding it from a big page-level wrapper
      // (`.slide`'s own box-shadow) onto every already-independently-
      // styled descendant (a badge that set its own background/
      // shrinkToFit) would give that badge a shadow it never had in the
      // source CSS.
      //
      // ALSO gated on `nodes.length === 1 || border` — the same "genuinely
      // one bounded box" signal cardGroupId's own `nodes.length > 1 &&
      // border` requirement uses (a multi-child card like `.card{border:1px
      // solid #ddd;border-radius:10px;box-shadow:...}` wrapping a heading +
      // two paragraphs legitimately needs shadow/rounded on its members too
      // — see the "User Card" test — so this can't just be nodes.length===1
      // outright). !styledByInnerWrapper alone doesn't catch a *plain,
      // otherwise totally unstyled* sibling (e.g. a bare `<h2>` sitting next
      // to other content inside `.slide{background:white;box-shadow:...}`,
      // no border of its own) — that heading has no border/backgroundColor/
      // etc of its own yet, so styledByInnerWrapper is false for it too, and
      // it would silently pick up `.slide`'s own shadow/rounded exactly like
      // a real single-child `.shadow{background:white;box-shadow:...}` div
      // does, with no way to tell the two apart from the node alone
      // afterward. Requiring either a single child (the shadow unambiguously
      // describes that one child) or a real border (the same strong,
      // deliberate "this is one bounded panel" signal cardGroupId already
      // trusts for the many-children case) excludes exactly the
      // `.slide`-style many-children/no-border wrapper this whole gate
      // exists to keep out.
      if (shadow && shadowRoundedTrusted && !styledByInnerWrapper && node.shadow === undefined) node.shadow = shadow;
      // Same rationale/gating as shadow immediately above.
      if (rounded && shadowRoundedTrusted && !styledByInnerWrapper && node.rounded === undefined) node.rounded = rounded;
      if (circular && shadowRoundedTrusted && !styledByInnerWrapper && node.circular === undefined) node.circular = circular;
      // Only genuinely inline-children types (paragraph/heading — see
      // INLINE_CHILDREN_NODE_TYPES's own note) get color inheritance:
      // wrapping `children` in an inline StyledSpanNode would corrupt a
      // quote/list/table/code/figure/section's own block children.
      if (INLINE_CHILDREN_NODE_TYPES.has(node.type) && inheritedColor) {
        node.children = [styledSpan({ color: inheritedColor }, node.children)];
      }
    }
  }
  if (extractPageBreakBefore(el.attrs?.style)) nodes.unshift({ type: 'pageBreak' });
  return nodes;
}

function extractPageBreakBefore(styleAttr) {
  const css = parseInlineStyleAttr(styleAttr);
  return css['page-break-before'] === 'always' || css['break-before'] === 'page' || css['break-before'] === 'always';
}

// Shared by <pre> and any other element preservesWhitespace recognizes as a
// code block (see mapElementCore's own note) — a dark-theme code block
// (`{background:#272822;color:#eee}`) is a very common real-world pattern,
// read directly off the element's own style same as any other
// explicit-declaration-on-the-element rule in this file. Falls back to the
// exporter's own default shading when the element declares neither.
//
// padding/margin are ALSO read directly here, not left to mapElement's own
// generic wrapper-bleed loop below (the mechanism every other box-styled
// node type relies on): that loop only bleeds padding/margin onto a node
// when `!paddingAlreadyOwned`, i.e. when nothing has already claimed the
// node via border/backgroundColor/etc — but this function itself just set
// backgroundColor on `code`'s own node moments before mapElement ever gets a
// turn, at construction time. That made paddingAlreadyOwned true for every
// dark-themed `<pre style="padding:12px">` before its own declared padding
// ever had a chance to attach, silently dropping it (and its margin) no
// matter what the source CSS said — the same class of "generic mechanism
// can't reach a node a special-cased constructor already claimed" bug
// BOX_STYLE_NODE_TYPES's own note describes for missing node types, just
// via early field assignment instead.
function codeFromStyledElement(el, value, lang) {
  const backgroundColor = extractBackgroundColor(el.attrs?.style);
  const color = extractRunStyle(el.attrs?.style)?.color;
  const padding = extractPaddingPt(el.attrs?.style);
  const margin = extractMarginPt(el.attrs?.style);
  return code(value, lang, backgroundColor, color, padding, margin);
}

const CLIPPING_OVERFLOW_VALUES = new Set(['auto', 'scroll', 'hidden']);

// A `<div>` styled like a browser scrollable box (`overflow:auto/scroll/
// hidden` on either axis, PLUS a declared height/max-height — overflow alone
// means nothing without a box to overflow OUT of) produces this private,
// non-public IDM marker instead of unwrapping like a plain div. It's resolved
// away entirely by resolve-scrollable-content.js (called from convert(),
// before normalizeDocument) according to the preserveScrollableContent
// option, so it never reaches an exporter — no exporter, and no other part
// of the IDM contract in idm/index.js, ever needs to know this type exists.
function mapScrollableContainer(el) {
  const style = el.attrs?.style;
  const overflowX = extractOverflowX(style);
  const overflowY = extractOverflowY(style);
  const isClipping = CLIPPING_OVERFLOW_VALUES.has(overflowX) || CLIPPING_OVERFLOW_VALUES.has(overflowY);
  const heightPx = extractHeightPx(style);
  const maxHeightPx = extractMaxHeightPx(style);
  if (!isClipping || (heightPx === undefined && maxHeightPx === undefined)) return undefined;
  return { type: '__scrollableContainer', heightPx, maxHeightPx, children: mapBlocks(el.children) };
}

// A block element's own inline style can carry character-level formatting
// too (e.g. <p style="color:#6b7280">), not just alignment — wrap the mapped
// inline content in a StyledSpanNode so it renders consistently with how
// <strong>/<em> content is formatted, via the same run-property accumulation.
//
// `backgroundColor` is deliberately stripped from that borrowed runStyle:
// every caller of this function (a <p>/<dt>/<dd>, or a table cell via
// mapCells) already gets — or, for paragraph/heading, is about to get from
// mapElement's own later pass over the very same element — its own
// dedicated block/cell-level `backgroundColor` field from this exact same
// style attribute (see mapCells' own `backgroundColor` extraction, and
// mapElement's `if (backgroundColor && node.backgroundColor === undefined)`
// assignment). Leaving it in here would extract the SAME declared background
// twice — once correctly as the element's own box/cell fill, and a second
// time as a run-level "highlight" wrapping its text — drawing a redundant
// second rect behind the text on every exporter that actually renders a
// styledSpan's own backgroundColor as a highlight. A genuine inline
// highlight (`<mark>`, a mid-sentence `<span style="background:...">`) never
// goes through this function at all — those map via their own dedicated
// case in mapInlineElementCore — so this exclusion can't affect real
// highlighting, only the duplicate.
// `fallbackRunStyle` is a <tr>'s own bold/italic/color/etc. (see
// extractRowStyle) — used only for a cell's plain phrasing content, so a
// `<tr style="font-weight:bold">` header row with unstyled `<td>`s still
// renders its text bold, matching a browser's inherited font-weight. The
// cell's own style always wins property-by-property over the row's.
function mapInlineWithStyle(el, fallbackRunStyle) {
  const children = mapInline(el.children);
  const ownRunStyle = extractRunStyle(el.attrs?.style);
  const runStyle = (fallbackRunStyle || ownRunStyle) ? { ...fallbackRunStyle, ...ownRunStyle } : undefined;
  if (runStyle) delete runStyle.backgroundColor;
  return runStyle && Object.keys(runStyle).length && children.length ? [styledSpan(runStyle, children)] : children;
}

// DROP_ELEMENTS/isHidden are already filtered by mapBlocks (this function's
// only caller path, via mapElement) before el ever reaches here — no guard
// needed on entry.
function mapElementCore(el) {
  switch (el.name) {
    // Headings
    case 'h1': case 'h2': case 'h3':
    case 'h4': case 'h5': case 'h6':
      return [heading(parseInt(el.name[1], 10), mapInlineWithStyle(el))];

    // Paragraph
    case 'p':
      return [paragraph(mapInlineWithStyle(el))];

    // A <button> is CSS inline-level but, unlike a real inline element,
    // essentially never sits mid-sentence in practice — it's a standalone
    // CTA (see TOP_LEVEL_BLOCK_TAGS above) — so it maps like <p>, but flagged
    // `shrinkToFit` so the DOCX exporter renders it as an auto-width box
    // (a single-cell autofit table) instead of a plain paragraph, which in
    // OOXML always shades/borders the full text column width with no
    // "size to content" option — the wrong shape for a button.
    case 'button': {
      const node = paragraph(mapInlineWithStyle(el));
      node.shrinkToFit = true;
      return [node];
    }

    // Lists
    case 'ul':
      return [list(false, mapListItems(el.children))];
    case 'ol':
      return [list(true, mapListItems(el.children), el.attrs['start'] ? parseInt(el.attrs['start'], 10) : undefined)];
    case 'li':
      return []; // handled by mapListItems

    // Definition list — <dt>/<dd> become separate paragraphs (previously
    // they fell through to the generic default case and collapsed into one
    // run with no separation at all, e.g. "TermDefinition"). <dt> is bolded,
    // the conventional way to visually distinguish a term from its definition.
    case 'dl':
      return mapBlocks(el.children);
    case 'dt':
      return [paragraph([strong(mapInline(el.children))])];
    case 'dd':
      return [paragraph(mapInline(el.children))];

    // Table
    case 'table':
      return [mapTable(el)];

    // Figure
    case 'figure': {
      const imgEl = findDescendant(el, 'img');
      if (!imgEl) return mapBlocks(el.children);
      const captionEl = findChild(el, 'figcaption');
      const caption = captionEl ? textOf(captionEl).trim() : undefined;
      const [w, h] = parseDimensionAttrs(imgEl);
      return [figure(image(imgEl.attrs['src'] ?? '', imgEl.attrs['alt'], undefined, w, h), caption)];
    }

    // Image
    case 'img':
      return [mapImgTag(el)];

    // SVG — dropped unless it's just a wrapper around one embedded raster
    // image (a common design-tool export pattern for logos/icons); vector
    // shapes/paths with no embedded raster have no representation we can
    // produce, so those still drop silently.
    case 'svg':
      return mapSvgTag(el);

    // Code blocks
    case 'pre': {
      const codeEl = findChild(el, 'code');
      const lang = codeEl ? langFromClass(codeEl.attrs['class']) : undefined;
      const value = codeEl ? textOf(codeEl) : textOf(el);
      return [codeFromStyledElement(el, value, lang)];
    }

    // Blockquote
    case 'blockquote':
      return [quote(mapBlocks(el.children))];

    // Sections
    case 'section':
    case 'article':
    case 'main':
      return [section(mapBlocks(el.children))];

    // Divs — just recurse, unless it's a flex/grid container (see
    // mapLayoutBlocks for why that needs different handling). Plain flow-
    // content wrappers (form/fieldset/address/details/summary/legend — see
    // CELL_BLOCK_TAGS's own note on why these are here) get the exact same
    // treatment: nothing about their own semantics (a form's submit
    // behavior, a details/summary disclosure toggle) has any document-export
    // equivalent, so recursing like a plain div is the correct fallback.
    case 'div':
    case 'header':
    case 'footer':
    case 'aside':
    case 'form':
    case 'fieldset':
    case 'address':
    case 'details':
    case 'summary':
    case 'legend':
    // <canvas>/<slot> — see CELL_BLOCK_TAGS's own note on why their content
    // (fallback / projected, respectively) needs the same block treatment.
    case 'canvas':
    case 'slot': {
      // A plain `<div>` styled to *look* like a code block (dark background,
      // monospace font, `white-space:pre`) but not using the semantic <pre>
      // tag — a common real-world pattern from CMS/editor output. Checked
      // before the flex/grid/plain-recursion fallbacks below: without this,
      // the div's own text (which DOES already carry its literal line
      // breaks/blank lines all the way from the source HTML — nothing
      // upstream collapses them) becomes an ordinary paragraph, and every
      // exporter's own word-wrapping (which treats all whitespace,
      // including newlines, as plain word separators) silently joins every
      // line into one run-on paragraph.
      if (preservesWhitespace(el.attrs?.style)) return [codeFromStyledElement(el, textOf(el).trim())];
      const displayMode = extractDisplayMode(el.attrs?.style);
      if (displayMode) return mapLayoutBlocks(el, displayMode);
      const scrollableContainer = mapScrollableContainer(el);
      return scrollableContainer ? [scrollableContainer] : mapBlocks(el.children);
    }

    // Thematic break
    case 'hr': {
      const cls = el.attrs['class'] ?? '';
      return [cls.includes('page-break') ? { type: 'pageBreak' } : { type: 'thematicBreak' }];
    }

    // A standalone <br> never actually reaches mapElementCore — it isn't in
    // TOP_LEVEL_BLOCK_TAGS, so mapBlocks always routes it into an inline run
    // (handled by mapInlineElementCore's own 'br' case) instead of here. Kept
    // as an explicit no-op for symmetry with that case, not dead code removal.
    case 'br':
      return [];

    // Inline elements in a block context — wrap in a paragraph
    default: {
      const inline = mapInline([el]);
      return inline.length ? [paragraph(inline)] : [];
    }
  }
}

/* -------------------------------------------------------------------------- */
/* List items                                                                  */
/* -------------------------------------------------------------------------- */

function mapListItems(nodes) {
  const items = [];
  for (const node of nodes) {
    if (node.type === 'text') continue;
    if (node.name !== 'li') continue;
    items.push(listItem(mapBlocks(node.children)));
  }
  return items;
}

/* -------------------------------------------------------------------------- */
/* Table                                                                       */
/* -------------------------------------------------------------------------- */

function mapTable(el) {
  const captionEl = findChild(el, 'caption');
  const caption = captionEl ? textOf(captionEl).trim() : undefined;
  const border = extractTableBorder(el.attrs?.style);
  // Neither has any other consumer today — only the XLSX exporter reads
  // them, to decide how many of the sheet's columns this table should span
  // and whether it should hug the left/right edge or center (see
  // dom-to-idm.js module doc / exporter/xlsx/document-sheet.js).
  const widthPercent = extractTableWidthPercent(el.attrs?.style);
  const marginAuto = extractMarginAuto(el.attrs?.style);

  const theadEl = findChild(el, 'thead');
  const tbodyEl = findChild(el, 'tbody') ?? el;
  const tfootEl = findChild(el, 'tfoot');

  const head = theadEl ? mapRows(theadEl.children, true) : undefined;
  // <tfoot> is a sibling of <tbody>, not nested inside it — its rows are
  // otherwise invisible to tbodyEl.children and get silently dropped.
  // Rendered as plain (non-header) rows appended after the body's own rows,
  // matching a browser's default rendering position for a table footer.
  const rows = [...mapRows(tbodyEl.children, false), ...(tfootEl ? mapRows(tfootEl.children, false) : [])];

  // If no explicit thead, treat the first row as head when all cells are <th>
  if (!head && rows.length > 0 && rows[0].cells.every((c) => c.header)) {
    const [first, ...rest] = rows;
    return { type: 'table', caption, border, widthPercent, marginAuto, head: [first], rows: rest };
  }

  return { type: 'table', caption, border, widthPercent, marginAuto, head, rows };
}

function mapRows(nodes, headerRow) {
  const rows = [];
  for (const node of nodes) {
    if (node.type !== 'element' || node.name !== 'tr') continue;
    const rowStyle = extractRowStyle(node.attrs?.style);
    const cells = mapCells(node.children, headerRow, rowStyle);
    if (cells.length) rows.push(tableRow(cells, rowHeightPx(node)));
  }
  return rows;
}

// A <tr>'s own background/border/font styling (e.g. the common
// `<tr style="background:#eee;border-bottom:2px solid #888">` header-row
// pattern, with the actual `<td>`s inside carrying no style of their own).
// `background-color`/`border`/`font-weight` don't CSS-inherit from parent to
// child, but a browser still visibly paints the row's background behind, and
// its border around, every cell in it — so this is read once per row and
// handed to mapCells as a fallback for whichever cells don't declare that
// property themselves; an explicit style on the cell always wins.
function extractRowStyle(styleAttr) {
  return {
    backgroundColor: extractBackgroundColor(styleAttr),
    borderTop: extractCellBorder(styleAttr, 'top'),
    borderBottom: extractCellBorder(styleAttr, 'bottom'),
    borderLeft: extractCellBorder(styleAttr, 'left'),
    borderRight: extractCellBorder(styleAttr, 'right'),
    runStyle: extractRunStyle(styleAttr),
  };
}

// A row's own height is rarely styled directly on <tr> — it's usually
// carried by one of its <td>s (e.g. a decorative colored strip cell styled
// height:4px) — so take the tallest explicit height declared by the row
// itself or any of its cells, matching how a browser sizes the row.
function rowHeightPx(trEl) {
  let max = extractHeightPx(trEl.attrs?.style);
  for (const cell of trEl.children) {
    if (cell.type !== 'element') continue;
    const h = extractHeightPx(cell.attrs?.style);
    if (h != null) max = max == null ? h : Math.max(max, h);
  }
  return max;
}

function mapCells(nodes, headerRow, rowStyle = {}) {
  const cells = [];
  for (const node of nodes) {
    if (node.type !== 'element') continue;
    if (node.name !== 'td' && node.name !== 'th') continue;
    const isHeader = headerRow || node.name === 'th';
    const colspan = node.attrs['colspan'] ? parseInt(node.attrs['colspan'], 10) : undefined;
    const rowspan = node.attrs['rowspan'] ? parseInt(node.attrs['rowspan'], 10) : undefined;
    const backgroundColor = extractBackgroundColor(node.attrs?.style) ?? rowStyle.backgroundColor;
    const widthPercent = extractWidthPercent(node);
    const borderTop = extractCellBorder(node.attrs?.style, 'top') ?? rowStyle.borderTop;
    const borderBottom = extractCellBorder(node.attrs?.style, 'bottom') ?? rowStyle.borderBottom;
    const borderLeft = extractCellBorder(node.attrs?.style, 'left') ?? rowStyle.borderLeft;
    const borderRight = extractCellBorder(node.attrs?.style, 'right') ?? rowStyle.borderRight;
    // An explicit `border:none`/`border-style:none` needs to actively
    // suppress a border the cell would otherwise inherit from the table's
    // own style (e.g. a data table's TableGrid) — see hasExplicitNoBorder's
    // own comment. `undefined` when not explicitly cancelled, so it doesn't
    // pollute every cell's serialized shape.
    const noBorder = hasExplicitNoBorder(node.attrs?.style) || undefined;
    const padding = extractPaddingPt(node.attrs?.style);
    // A <th> is browser-default center-aligned (with <td> left-aligned) even
    // with no CSS at all — the same "an established convention counts as an
    // explicit declaration" reasoning already used for <mark>/<sub>/<u>
    // elsewhere in this file. An explicit `text-align` on the cell itself
    // always wins over that default, same as it would in a browser.
    const align = extractAlign(node.attrs?.style) ?? (isHeader ? 'center' : undefined);
    const verticalAlign = extractVerticalAlign(node.attrs?.style);
    const direction = extractDirection(node.attrs?.style);
    const writingMode = extractWritingMode(node.attrs?.style);
    const rotateDeg = extractRotationDeg(node.attrs?.style);
    // A cell's own font styling — parallel to how backgroundColor above is
    // read from the same style attribute for the cell's own shading,
    // independent of the run-level styledSpan the same CSS also produces
    // around the cell's flattened inline content (see mapInlineWithStyle).
    // Excel has one font per cell (not per run) when the cell's content
    // resolves to a single run, so XLSX reads these directly rather than
    // inspecting rendered children — DOCX/CSV/etc. don't read these new
    // fields, so this is purely additive.
    const ownRunStyle = extractRunStyle(node.attrs?.style);
    const runStyle = (rowStyle.runStyle || ownRunStyle) ? { ...rowStyle.runStyle, ...ownRunStyle } : undefined;
    const color = runStyle?.color;
    const bold = runStyle?.bold;
    const italic = runStyle?.italic;
    const underline = runStyle?.underline;
    const strikethrough = runStyle?.strikethrough;
    const fontSizePt = runStyle?.fontSizePt;
    // A <td>/<th> itself styled display:none/visibility:hidden must still keep
    // its slot (dropping it outright would corrupt colspan/rowspan geometry
    // for every other cell in the table) — only its content is suppressed.
    // A cell with only phrasing content (text, <strong>, <a>, ...) is mapped inline,
    // same as before. A cell with block-level content (a nested <table>, <p>, <div>
    // wrapping blocks, ...) must go through mapBlocks — otherwise mapInline's fallback
    // for unrecognized tags silently unwraps nested rows/cells into one flat text blob.
    const content = isHidden(node.attrs?.style) ? []
      : hasBlockContent(node.children) ? mapBlocks(node.children) : mapInlineWithStyle(node, rowStyle.runStyle);
    cells.push(tableCell(content, {
      header: isHeader, colspan, rowspan, backgroundColor, widthPercent,
      borderTop, borderBottom, borderLeft, borderRight, noBorder, padding, align, verticalAlign,
      color, bold, italic, underline, strikethrough, fontSizePt,
      direction, writingMode, rotateDeg,
    }));
  }
  return cells;
}

function hasBlockContent(nodes) {
  return nodes.some((n) => n.type === 'element' && CELL_BLOCK_TAGS.has(n.name));
}

// <td width="52%"> or <td style="width:52%">. Pixel/other units aren't
// handled — a percentage is the only width form that translates cleanly to
// a proportional share of a fixed page content width in the output formats.
function extractWidthPercent(node) {
  const raw = node.attrs?.['width'] ?? parseInlineStyleAttr(node.attrs?.style)['width'];
  if (!raw) return undefined;
  const m = /^([\d.]+)%$/.exec(String(raw).trim());
  return m ? Number(m[1]) : undefined;
}

/* -------------------------------------------------------------------------- */
/* Inline mapping                                                              */
/* -------------------------------------------------------------------------- */

function mapInline(nodes) {
  const out = [];
  for (const node of nodes) {
    if (node.type === 'text') {
      if (node.value) out.push(text(node.value));
      continue;
    }
    out.push(...mapInlineElement(node));
  }
  return out;
}

// A plain-looking <span>/<b>/<a>/... may still carry its own explicit inline
// style (very common from CMS/editor/email-client HTML that uses CSS instead
// of semantic tags for bold/color — exactly the "GENERATED FOR" labels case).
// Wrap whatever the tag-specific mapping below produced in a StyledSpanNode
// so it renders with that formatting too, regardless of which tag carried it.
function mapInlineElement(el) {
  if (DROP_ELEMENTS.has(el.name) || isHidden(el.attrs?.style)) return [];
  const nodes = mapInlineElementCore(el);
  const runStyle = extractRunStyle(el.attrs?.style);
  return runStyle && nodes.length ? [styledSpan(runStyle, nodes)] : nodes;
}

// DROP_ELEMENTS/isHidden are already filtered by mapInlineElement (this
// function's only caller) just above — no guard needed on entry.
function mapInlineElementCore(el) {
  switch (el.name) {
    case 'strong': case 'b':
      return [strong(mapInline(el.children))];
    case 'em': case 'i':
      return [emphasis(mapInline(el.children))];
    case 'code':
      return [codeSpan(textOf(el))];
    case 'a': {
      const href = el.attrs['href'] ?? '#';
      const title = el.attrs['title'];
      return [link(href, mapInline(el.children), title)];
    }
    case 'br':
      return [{ type: 'lineBreak' }];
    case 'img':
      return [mapImgTag(el)];
    case 'svg':
      return mapSvgTag(el);
    // Underline/strikethrough are the same run-level styles the CSS
    // text-decoration path already produces (see extractRunStyle) — an
    // explicit <u>/<s>/<del>/<ins> tag carries the same formatting the same
    // way <strong> carries bold, so it reuses the same StyledSpanNode.
    case 'u': case 'ins':
      return [styledSpan({ underline: true }, mapInline(el.children))];
    case 's': case 'strike': case 'del':
      return [styledSpan({ strikethrough: true }, mapInline(el.children))];
    // <mark> is a highlight — background-color, same as inline CSS background.
    case 'mark':
      return [styledSpan({ backgroundColor: 'FFFF00' }, mapInline(el.children))];
    case 'sub':
      return [styledSpan({ verticalAlign: 'subscript' }, mapInline(el.children))];
    case 'sup':
      return [styledSpan({ verticalAlign: 'superscript' }, mapInline(el.children))];
    // <kbd> (keyboard input) conventionally renders like code — monospace.
    case 'kbd':
      return [codeSpan(textOf(el))];
    // <var>/<cite> (a variable name / a work's title) are conventionally italic.
    case 'var': case 'cite':
      return [styledSpan({ italic: true }, mapInline(el.children))];
    // <q> is defined by HTML to render surrounded by quotation marks — this
    // is the tag's literal spec behavior, not an inferred style.
    case 'q':
      return [text('"'), ...mapInline(el.children), text('"')];
    // A checkbox's checked state is real, explicit content (the same way an
    // <img alt> or a <select>'s chosen <option> is) — represented with the
    // conventional Unicode checkbox glyphs. Other input types (text, email,
    // ...) have no content to show when empty, so they're left as nothing
    // rather than fabricating a placeholder.
    case 'input': {
      const type = (el.attrs['type'] ?? 'text').toLowerCase();
      if (type === 'checkbox' || type === 'radio') {
        return [text(el.attrs['checked'] !== undefined ? '☑' : '☐')];
      }
      // A browser never shows a hidden field's value at all, and a password
      // field shows masking dots rather than the real value — its actual
      // *visible* content is the mask, not a placeholder to fill in.
      // Anything else (text/email/number/date/tel/url/search/...) shows its
      // value verbatim, the same "real, explicit content" reasoning already
      // applied to checkbox/radio/select above.
      if (type === 'hidden') return [];
      const value = el.attrs['value'];
      if (!value) return [];
      if (type === 'password') return [styledSpan({ formField: true }, [text('•'.repeat(value.length))])];
      return [styledSpan({ formField: true }, [text(value)])];
    }
    // A <textarea>'s content, same "formField" treatment as <input> above —
    // DOCX renders it as a real fillable content-control box; every other
    // exporter (which doesn't understand that style key) just sees plain text.
    case 'textarea':
      return [styledSpan({ formField: true }, [text(textOf(el))])];
    // A <select>'s explicitly chosen value (its selected <option>, or the
    // first one — the same default a browser applies when none is marked
    // selected) is its real content; showing every <option> concatenated
    // together (the previous default-case behavior) produced unreadable
    // runs like "OneTwo" with no indication of the actual selection.
    case 'select': {
      const options = el.children.filter((c) => c.type === 'element' && c.name === 'option');
      const chosen = options.find((o) => o.attrs['selected'] !== undefined) ?? options[0];
      return chosen ? [styledSpan({ formField: true }, [text(textOf(chosen))])] : [];
    }
    case 'span':
    case 'label':
    case 'abbr':
    case 'dfn':
    case 'small':
      return mapInline(el.children);
    default:
      return mapInline(el.children);
  }
}

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

function textOf(el) {
  let out = '';
  for (const child of el.children) {
    if (child.type === 'text') out += child.value;
    else out += textOf(child);
  }
  return out;
}

// Shared by mapElementCore's and mapInlineElementCore's identical 'img'
// cases (block vs inline context maps an <img> the same way either time).
// A real browser applies CSS `width`/`height` (e.g. `style="width:80px"`)
// OVER the bare HTML attribute of the same name — the attribute is only a
// presentational hint with the lowest possible specificity — so
// extractWidthPx/extractHeightPx (already used for this exact CSS
// precedence elsewhere in this file) win whenever the source declared
// both; the HTML attribute is only the fallback for an <img> with no CSS
// size at all. Without this, an `<img style="width:80px;height:80px">`
// whose real decoded pixel dimensions are tiny or huge (a 1x1 tracking
// pixel, a 4000px camera photo) rendered at its own real intrinsic size in
// every exporter instead of the size the page actually displays it at.
function mapImgTag(el) {
  const [attrW, attrH] = parseDimensionAttrs(el);
  const w = extractWidthPx(el.attrs?.style) ?? attrW;
  const h = extractHeightPx(el.attrs?.style) ?? attrH;
  return image(el.attrs['src'] ?? '', el.attrs['alt'], undefined, w, h);
}

// Shared by mapElementCore's and mapInlineElementCore's identical 'svg'
// cases — an array (possibly empty) so both callers can `return mapSvgTag(el)` directly.
// Two SVG patterns are recognized, in priority order: an embedded raster
// (the older, still-common design-tool-export case, handled first since an
// <svg> that's purely a wrapper around one <image> has no separate vector
// content worth extracting alongside it) — else genuine vector primitives
// (rect/circle/ellipse/line/polygon/polyline). Anything else (a `<path>`,
// gradients, `<use>`, `<text>`, a transformed `<g>`) has no representation
// this project produces and is silently dropped, same as it always was.
function mapSvgTag(el) {
  const src = extractSvgImageSrc(el);
  if (src) {
    const [w, h] = parseDimensionAttrs(el);
    return [image(src, svgAlt(el), undefined, w, h)];
  }
  const shapes = extractSvgShapes(el);
  return shapes ? [shapes] : [];
}

// SVG's own <image> element (distinct from HTML's <img>) commonly wraps a
// base64 raster image via xlink:href/href — the actual pattern found in
// design-tool-exported logos. Only that direct case is handled: a real
// vector shape (paths/rects with no embedded raster) has nothing to extract.
function extractSvgImageSrc(el) {
  const imageEl = findDescendant(el, 'image');
  if (!imageEl) return undefined;
  const href = imageEl.attrs['xlink:href'] ?? imageEl.attrs['href'];
  return href && href.startsWith('data:image/') ? href : undefined;
}

/* -------------------------------------------------------------------------- */
/* SVG vector shapes                                                           */
/* -------------------------------------------------------------------------- */
// Recognizes exactly the "basic shapes" subset of SVG — rect, circle,
// ellipse, line, polygon, polyline, each with fill/stroke/stroke-width — the
// common case for icons, logos, and simple diagrams. Deliberately excludes
// <path> (arbitrary Bézier/arc path data), gradients-on-shapes, <text>,
// <use>, <defs>/<symbol>, filters/masks, and any transformed <g> (a
// transform-less <g> is still walked into — pure semantic grouping with no
// coordinate shift is common even in simple icon SVGs and costs nothing to
// support) — each is a real, substantially larger undertaking on its own,
// deliberately left for a later pass rather than guessed at.

const SVG_SHAPE_EXTRACTORS = {
  rect: extractRectShape,
  circle: extractCircleShape,
  ellipse: extractEllipseShape,
  line: extractLineShape,
  polygon: (el) => extractPolyShape(el, 'polygon'),
  polyline: (el) => extractPolyShape(el, 'polyline'),
};

const DEFAULT_SVG_WIDTH_PX = 300;
const DEFAULT_SVG_HEIGHT_PX = 150;

function extractSvgShapes(el) {
  const shapeElements = [];
  collectSvgShapeElements(el, shapeElements);
  if (shapeElements.length === 0) return undefined;

  const shapes = [];
  for (const shapeEl of shapeElements) {
    const shape = SVG_SHAPE_EXTRACTORS[shapeEl.name](shapeEl);
    if (shape) shapes.push(shape);
  }
  if (shapes.length === 0) return undefined;

  const [declaredWidthPx, declaredHeightPx] = parseDimensionAttrs(el);
  // The tokenizer lowercases every attribute name (HTML attributes are
  // case-insensitive), so SVG's own camelCase `viewBox` is stored as
  // `viewbox` in `el.attrs` — same reason `xlink:href` above needed no such
  // adjustment (it was already all-lowercase to begin with).
  const viewBox = parseViewBoxAttr(el.attrs?.viewbox);
  const widthPx = declaredWidthPx ?? viewBox?.width ?? DEFAULT_SVG_WIDTH_PX;
  const heightPx = declaredHeightPx ?? viewBox?.height ?? DEFAULT_SVG_HEIGHT_PX;

  // A `viewBox` describes the coordinate space shapes are actually drawn
  // in, which can differ from the declared display width/height (e.g.
  // `viewBox="0 0 24 24"` on a `width="48"` icon) — every shape's own
  // coordinates are rescaled here, once, into the final `widthPx`/
  // `heightPx` px box, so no exporter needs to think about viewBox at all.
  if (viewBox) {
    const scaleX = widthPx / viewBox.width;
    const scaleY = heightPx / viewBox.height;
    if (viewBox.minX !== 0 || viewBox.minY !== 0 || scaleX !== 1 || scaleY !== 1) {
      for (const shape of shapes) rescaleShape(shape, viewBox.minX, viewBox.minY, scaleX, scaleY);
    }
  }

  return svgShapes(shapes, widthPx, heightPx);
}

// Walks direct children, plus one level into any transform-less <g> (pure
// grouping — see extractSvgShapes's own note on why a transformed <g> is
// skipped instead). Recurses through nested transform-less <g>s too.
function collectSvgShapeElements(el, out) {
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    if (child.name === 'g' && !child.attrs?.['transform']) {
      collectSvgShapeElements(child, out);
      continue;
    }
    if (SVG_SHAPE_EXTRACTORS[child.name]) out.push(child);
  }
}

function svgNumber(value, fallback = 0) {
  const n = parseFloat(value);
  return Number.isFinite(n) ? n : fallback;
}

// SVG's own presentation-attribute rules: `style="..."` always wins over a
// same-named bare attribute (e.g. `<rect fill="red" style="fill:blue">` is
// blue), exactly like CSS specificity elsewhere in this file. `fill`
// defaults to black when declared nowhere at all — SVG's own real default,
// easy to get backwards since every OTHER shape property here defaults to
// "off" — `stroke` defaults to none. An explicit `fill="none"`/
// `stroke="none"` is the same "off" as leaving it undeclared.
function resolveSvgColor(css, attrs, prop, defaultHex) {
  const raw = css[prop] ?? attrs[prop];
  if (raw === undefined) return defaultHex;
  if (raw.trim().toLowerCase() === 'none') return undefined;
  return normalizeColor(raw) ?? defaultHex;
}

function svgPresentation(el) {
  const css = parseInlineStyleAttr(el.attrs?.style);
  const attrs = el.attrs ?? {};
  const fill = resolveSvgColor(css, attrs, 'fill', '000000');
  const stroke = resolveSvgColor(css, attrs, 'stroke', undefined);
  // stroke-width only matters (and only gets carried) when there's an
  // actual stroke to draw — SVG's own default stroke-width (1 user unit)
  // applies whenever a stroke color is set but no width is declared.
  const strokeWidthPx = stroke ? svgNumber(css['stroke-width'] ?? attrs['stroke-width'], 1) : undefined;
  return { fill, stroke, strokeWidthPx };
}

function extractRectShape(el) {
  const x = svgNumber(el.attrs?.x);
  const y = svgNumber(el.attrs?.y);
  const width = svgNumber(el.attrs?.width);
  const height = svgNumber(el.attrs?.height);
  if (width <= 0 || height <= 0) return undefined;
  return { kind: 'rect', x, y, width, height, ...svgPresentation(el) };
}

function extractCircleShape(el) {
  const cx = svgNumber(el.attrs?.cx);
  const cy = svgNumber(el.attrs?.cy);
  const r = svgNumber(el.attrs?.r);
  if (r <= 0) return undefined;
  return { kind: 'circle', cx, cy, r, ...svgPresentation(el) };
}

function extractEllipseShape(el) {
  const cx = svgNumber(el.attrs?.cx);
  const cy = svgNumber(el.attrs?.cy);
  const rx = svgNumber(el.attrs?.rx);
  const ry = svgNumber(el.attrs?.ry);
  if (rx <= 0 || ry <= 0) return undefined;
  return { kind: 'ellipse', cx, cy, rx, ry, ...svgPresentation(el) };
}

// A <line> has no fill concept at all (SVG spec) — only its own stroke, and
// with no stroke declared it draws literally nothing, so it's dropped here
// rather than carried as an invisible no-op shape.
function extractLineShape(el) {
  const x1 = svgNumber(el.attrs?.x1);
  const y1 = svgNumber(el.attrs?.y1);
  const x2 = svgNumber(el.attrs?.x2);
  const y2 = svgNumber(el.attrs?.y2);
  const { stroke, strokeWidthPx } = svgPresentation(el);
  return stroke ? { kind: 'line', x1, y1, x2, y2, stroke, strokeWidthPx } : undefined;
}

function parsePointsAttr(value) {
  const numbers = (value ?? '').trim().split(/[\s,]+/).filter(Boolean).map(Number).filter(Number.isFinite);
  const points = [];
  for (let i = 0; i + 1 < numbers.length; i += 2) points.push([numbers[i], numbers[i + 1]]);
  return points;
}

function extractPolyShape(el, kind) {
  const points = parsePointsAttr(el.attrs?.points);
  return points.length >= 2 ? { kind, points, ...svgPresentation(el) } : undefined;
}

function parseViewBoxAttr(value) {
  if (!value) return undefined;
  const parts = value.trim().split(/[\s,]+/).map(Number);
  if (parts.length !== 4 || parts.some((n) => !Number.isFinite(n)) || parts[2] <= 0 || parts[3] <= 0) return undefined;
  const [minX, minY, width, height] = parts;
  return { minX, minY, width, height };
}

// Mutates `shape`'s own coordinate fields in place from viewBox space into
// final px space — called once per shape immediately after extraction,
// before any exporter ever sees it. A circle can't stay a true circle under
// non-uniform X/Y scaling; averaging the two axis scales for its radius is
// the same "reasonable approximation, not a full layout engine" tradeoff
// this project already accepts elsewhere (e.g. list-item indent steps).
function rescaleShape(shape, minX, minY, scaleX, scaleY) {
  const avgScale = (scaleX + scaleY) / 2;
  switch (shape.kind) {
    case 'rect':
      shape.x = (shape.x - minX) * scaleX;
      shape.y = (shape.y - minY) * scaleY;
      shape.width *= scaleX;
      shape.height *= scaleY;
      break;
    case 'circle':
      shape.cx = (shape.cx - minX) * scaleX;
      shape.cy = (shape.cy - minY) * scaleY;
      shape.r *= avgScale;
      break;
    case 'ellipse':
      shape.cx = (shape.cx - minX) * scaleX;
      shape.cy = (shape.cy - minY) * scaleY;
      shape.rx *= scaleX;
      shape.ry *= scaleY;
      break;
    case 'line':
      shape.x1 = (shape.x1 - minX) * scaleX;
      shape.y1 = (shape.y1 - minY) * scaleY;
      shape.x2 = (shape.x2 - minX) * scaleX;
      shape.y2 = (shape.y2 - minY) * scaleY;
      break;
    case 'polygon':
    case 'polyline':
      shape.points = shape.points.map(([x, y]) => [(x - minX) * scaleX, (y - minY) * scaleY]);
      break;
  }
  if (shape.strokeWidthPx != null) shape.strokeWidthPx *= avgScale;
}

function svgAlt(el) {
  return el.attrs['aria-label'] ?? el.attrs['title'];
}

// HTML width/height attributes are unitless numbers meaning CSS px (e.g.
// <img width="209" height="56">, <svg width="209" height="56">) — the actual
// intrinsic size declared by the source markup, not inferred from decoding
// image bytes.
function parseDimensionAttrs(el) {
  return [toPx(el.attrs?.['width']), toPx(el.attrs?.['height'])];
}

function toPx(value) {
  if (!value) return undefined;
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? n : undefined;
}

function findDescendant(el, name) {
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    if (child.name === name) return child;
    const found = findDescendant(child, name);
    if (found) return found;
  }
  return undefined;
}

function langFromClass(cls) {
  if (!cls) return undefined;
  const m = cls.match(/language-(\S+)/);
  return m?.[1];
}
