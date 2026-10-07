/**
 * Groups the IDM document into slides and renders each one's
 * ppt/slides/slideN.xml. Two concerns live here that don't belong in
 * shapes.js: deciding *where* slide boundaries fall (PP2), and the coarse
 * vertical layout that hands each shape its box (title strip, then a body
 * placeholder, then any images/tables stacked below it).
 */
import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';
import { NS_A, NS_P, NS_R, EMU_PER_IN } from './layout.js';
import { renderBodyBlockParagraphs, renderPictureShape, renderTableShape, renderStyledBoxShape, renderCardShape, renderSvgShapesShape, renderThematicBreakShape, estimateCardHeightEmu, estimateCardBlockHeightPt, CARD_BLOCK_SPACING_PT, isStyledBoxNode } from './shapes.js';

const EMU_PER_PT = 12700; // matches the constant already used in shapes.js

// `thematicBreak` (a bare `<hr>`) is included here — see shapes.js's own
// renderThematicBreakShape note on why: without it, a standalone `<hr>`
// falls through to a plain "text" unit with no renderable content,
// silently claiming (and leaving empty) the slide's one real body
// placeholder.
const MEDIA_TYPES = new Set(['image', 'figure', 'table', 'svgShapes', 'thematicBreak']);

// A top-level <section>/<article>/<main> becomes its own slide, verbatim —
// the literal PP2 rule. But most real-world HTML (a CMS/editor export, LLM
// output, a scraped page) either never uses those tags, or — just as
// commonly — uses ONE incidentally, deep inside an otherwise flat document.
// Treating "any section present" as an all-or-nothing switch to
// section-only slides (this exporter's first pass) silently dropped every
// other top-level block in exactly that mixed case: a real dev fixture with
// 25 flat top-level blocks plus one trailing <section> collapsed to a single
// slide. So a <section> is instead a slide boundary encountered *inline*
// while walking the document — the surrounding flat content still splits on
// every top-level h1/h2 exactly as it would with no sections at all.
export function collectSlideGroups(document) {
  const children = document.children ?? [];
  const groups = [];
  let openFlowGroup = null; // the trailing non-section group currently being appended to, if any

  for (const node of children) {
    if (node.type === 'section') {
      groups.push(node.children ?? []);
      openFlowGroup = null; // content right after a section starts fresh, not merged into it
      continue;
    }
    const startsSlide = node.type === 'heading' && node.level <= 2;
    if (startsSlide || !openFlowGroup) {
      openFlowGroup = [];
      groups.push(openFlowGroup);
    }
    openFlowGroup.push(node);
  }
  return groups.length ? groups : [[]];
}

// Only a group's very first node becomes the title placeholder — a second
// h1/h2 later in the same slide (or a <section> that doesn't open with a
// heading at all) just renders inline in the body instead of being promoted;
// there's only one title placeholder per slide to give it to. A heading that
// belongs to a cardGroup (e.g. a `.hero{background:...}` banner's own <h1>
// sitting alongside a subtitle <p>, both meant to render inside one shared
// colored panel — see dom-to-idm.js's mapElement) must NOT be pulled out
// here even when it's the slide's first node: promoting it to its own
// separately-styled title placeholder splits it away from classifyContentUnits's
// cardGroup handling entirely, leaving only the rest of the group inside the
// real shape and the heading floating outside it with just a plain
// run-highlight instead of sharing the panel.
//
// A LONE self-styled banner heading (isStyledBoxNode — e.g. this project's
// own `<h2 style="background:#1F4E79;color:white;padding:8px">Invoice
// Items</h2>`, with no wrapping cardGroup div at all) needs the exact same
// exclusion for the exact same reason: classifyContentUnits' own
// isStyledBoxNode check (see its call site further down) is what turns a
// heading like this into a real full-width colored CARD shape — the title
// placeholder's own plain run-highlight bg handling only ever covers the
// glyphs themselves, not the placeholder's full width, so promoting it here
// silently downgraded a full-width banner into a tight highlighted swatch
// sized to its own short text, whenever it happened to land as a new
// slide's first node.
function splitTitleAndBody(group) {
  const [first, ...rest] = group;
  if (first?.type === 'heading' && first.level <= 2 && first.cardGroup === undefined && !isStyledBoxNode(first)) return { title: first, rest };
  return { title: undefined, rest: group };
}

// Images/figures/tables aren't text-flow content — DrawingML has no inline
// picture/table concept the way a paragraph run does, so they always need
// their own positioned shape. A short bordered/filled block (isStyledBoxNode
// — see shapes.js) is text-flow-shaped but still needs its own shape rather
// than stretching across the body placeholder — a shrink-wrapped badge/pill
// when it has a real "shrink to content" signal (shrinkToFit), otherwise a
// full-width single-node "card" (the common full-width gradient button/
// banner case) — and a run of consecutive nodes sharing one `cardGroup` (a
// bordered/filled <div> wrapping several blocks — see dom-to-idm.js's
// mapElement) needs one shared background shape behind all of them rather
// than either of those.
//
// Returns an ORDERED list of `{kind, ...}` units in the same relative order
// the source document had — text/blocks/media used to be split into three
// separate buckets and rendered in a fixed sequence (all blocks, then all
// text in one shared placeholder, then all media divided evenly across
// whatever space was left), which silently discarded the source's own
// interleaving (a heading immediately followed by its own image ended up
// nowhere near it) and gave every image the same fixed slice regardless of
// its actual size — the "6 images sharing one thumbnail-sized strip" bug.
// Consecutive plain text-flow nodes still merge into one unit (so 3
// paragraphs in a row become one body-placeholder shape, not three), but a
// unit boundary now also opens whenever the node's OWN kind changes, so the
// original order survives.
//
// layoutGroup membership is checked before cardGroup/isStyledBoxNode: a
// flex/grid item is very often ALSO a badge or card on its own (a colored
// `.box`, a bordered `.card`) — mapLayoutBlocks stamps layoutGroup onto
// those same nodes without disturbing their own cardGroup/backgroundColor,
// so checking it first pulls the *whole row* out as one positioned unit
// instead of peeling each item off individually into the plain vertical
// stack. cardGroup is checked next, then isStyledBoxNode: mapElement
// stamps the *same* border/background onto every member of a card, so a
// short paragraph inside one (e.g. a one-line "Role: Software Engineer")
// would otherwise also satisfy the badge heuristic on its own and get
// peeled off into its own separate shape instead of staying part of the card.
function classifyContentUnits(nodes) {
  const units = [];
  const pushText = (node) => {
    const last = units[units.length - 1];
    if (last?.kind === 'text') last.nodes.push(node);
    else units.push({ kind: 'text', nodes: [node] });
  };

  let i = 0;
  while (i < nodes.length) {
    const node = nodes[i];
    if (MEDIA_TYPES.has(node.type)) {
      units.push({ kind: 'media', node });
      i++;
    } else if (node.layoutGroup !== undefined) {
      const layoutGroup = node.layoutGroup;
      const cells = [];
      let cell = [node];
      let j = i + 1;
      while (j < nodes.length && nodes[j].layoutGroup === layoutGroup) {
        if (nodes[j].layoutItemIndex !== cell[0].layoutItemIndex) {
          cells.push(cell);
          cell = [];
        }
        cell.push(nodes[j]);
        j++;
      }
      cells.push(cell);
      units.push({ kind: 'block', block: { kind: 'layout', layoutGroup, cells } });
      i = j;
    } else if (node.cardGroup !== undefined) {
      const group = [node];
      let j = i + 1;
      while (j < nodes.length && nodes[j].cardGroup === node.cardGroup) group.push(nodes[j++]);
      units.push({ kind: 'block', block: { kind: 'card', nodes: group } });
      i = j;
    } else if (node.type === 'quote' && node.backgroundColor &&
      (node.border || node.borderTop || node.borderRight || node.borderBottom || node.borderLeft)) {
      // Requires a real border alongside backgroundColor — not backgroundColor
      // alone — for the same reason isStyledBoxNode requires it for a
      // paragraph/heading (see that function's own note): a quote's
      // backgroundColor bleeds down the same plain "if unset" way a
      // paragraph/heading's does (see dom-to-idm.js's quote branch — it's
      // deliberately NOT gated the way padding/margin are, so plain quote
      // text can still visually sit on an unrelated backdrop via the
      // run-highlight fallback below when nothing stronger applies). A bare,
      // completely unstyled `<blockquote>` sitting inside a big multi-child
      // `.page{background:#fff;box-shadow:...}` wrapper picked up that
      // backgroundColor exactly like any other paragraph does, and — before
      // this check — that alone was enough to wrongly promote it into its
      // own separate (here, invisible-because-white) shape. A real
      // `blockquote{border-left:...}` accent stripe, the actual common
      // pattern this promotion exists for, still clears this easily.
      //
      // A blockquote's own background (e.g. `blockquote{background:#eef4ff}`)
      // used to fall through to renderBodyBlockParagraphs's plain run-
      // highlight approximation — fine for a short paragraph/heading (see
      // isStyledBoxNode's own BADGE_MAX_CHARS reasoning), but a blockquote's
      // text is almost always long enough to wrap onto 2+ lines, and a
      // highlight only ever covers the actual glyphs on each wrapped line —
      // it visibly does NOT span the full text-column width the way DOCX's
      // real paragraph shading (w:shd) does, so a multi-line quote reads as a
      // ragged, inconsistent highlight instead of one clean box. Promoting it
      // to a real card shape (same as a bordered/gradient div) gives it a
      // proper full-width/full-height solid fill instead, matching DOCX.
      units.push({ kind: 'block', block: { kind: 'card', nodes: [node] } });
      i++;
    } else if (isStyledBoxNode(node)) {
      // `shrinkToFit` (display:inline-block, or an explicit width — see
      // dom-to-idm.js's mapElement) is what actually distinguishes a real
      // shrink-wrapped pill/badge from an ordinary block-level `<div>` that
      // just happens to have short text — e.g. a full-width gradient
      // button/banner. Without that signal, it renders as its own full-width
      // "card" (reusing renderCardShape with a single-node group) instead of
      // a shrink-wrapped badge.
      if (node.shrinkToFit) units.push({ kind: 'block', block: { kind: 'box', node } });
      else units.push({ kind: 'block', block: { kind: 'card', nodes: [node] } });
      i++;
    } else {
      pushText(node);
      i++;
    }
  }
  return units;
}

// `phType` is optional now: a slide can have more than one run of plain
// text-flow content once text/media/blocks are rendered in original
// document order (see classifyContentUnits) — e.g. "paragraph, image,
// paragraph, image" — but a slide layout only ever defines ONE real body
// placeholder (idx 1), so only the FIRST text run can claim it (see
// buildSlideXml's own bodyPlaceholderClaimed flag). Every text run after
// that renders as a plain freestanding text box instead: no `<p:ph>` at
// all, and — unlike a placeholder, which inherits its geometry from the
// slide layout when none is given — an explicit `<a:prstGeom prst="rect">`
// so PowerPoint has an actual shape to hang the text body on.
function shapeXml(tag, { shapeId, name, phType, phIdx, box }, bodyXml) {
  const b = new XmlBuilder();
  b.el(tag);
  b.el('p:nvSpPr');
  b.self('p:cNvPr', { id: shapeId, name });
  b.self('p:cNvSpPr');
  b.el('p:nvPr');
  if (phType) b.self('p:ph', { type: phType, idx: phIdx });
  b.up();
  b.up();
  b.el('p:spPr');
  b.el('a:xfrm');
  b.self('a:off', { x: box.x, y: box.y });
  b.self('a:ext', { cx: box.cx, cy: box.cy });
  b.up();
  if (!phType) {
    b.el('a:prstGeom', { prst: 'rect' });
    b.self('a:avLst');
    b.up();
  }
  b.up();
  b.el('p:txBody');
  b.self('a:bodyPr');
  b.self('a:lstStyle');
  b.raw(bodyXml);
  b.up();
  b.up();
  return b.end();
}

// `nextShapeId` is a callback, not a plain number — most media types here
// draw exactly one shape (one id), but an SVG draws one shape per vector
// primitive, and a table draws one extra floating picture per cell that
// holds an image (a native `a:tc` can't nest a `p:pic` — see
// findCellImageNode's own doc comment in shapes.js) — a variable count only
// each function itself knows ahead of time.
function renderMediaShape(node, rctx, box, nextShapeId) {
  if (node.type === 'image') return renderPictureShape(node, rctx, box, nextShapeId());
  if (node.type === 'figure') return renderPictureShape(node.image, rctx, box, nextShapeId());
  if (node.type === 'table') return renderTableShape(node, rctx, box, nextShapeId(), nextShapeId);
  if (node.type === 'svgShapes') return renderSvgShapesShape(node, rctx, box, nextShapeId);
  if (node.type === 'thematicBreak') return renderThematicBreakShape(box, nextShapeId());
  return { xml: '', cy: 0 };
}

/**
 * `rctx` is per-slide: `{ relationships, assetInfoByAssetId, images }` —
 * see shapes.js's resolveImageRel for why the image cache can't be shared
 * across slides the way DOCX's single RelationshipManager is.
 */
// Badges and cards both stack one per line, top to bottom — real CSS could
// flow several inline-block badges side by side, but without a
// text-measurement pass to know how many fit per row, stacking vertically is
// the safe choice that never overlaps two into the same space. A `layout`
// block (see renderLayoutBlock) is the one exception: `display:flex`/`grid`
// is an explicit, unambiguous signal to lay its own cells out side by side,
// unlike the "maybe inline-block" guesswork a bare badge/card has to accept.
const BOX_GAP_EMU = Math.round(EMU_PER_IN * 0.1);

// Renders one `display:flex`/`display:grid` container's cells as real
// side-by-side (flex) or wrapped (grid) shapes — every cell gets an equal
// share of `box.cx`, split into `columns` columns and wrapping into more
// rows once `cells.length` exceeds that. `columns` defaults to one row of
// N columns (flex's own single-row semantics, and grid's own fallback when
// `grid-template-columns` didn't parse to a track count). Every cell in the
// group shares one uniform row height — the tallest cell's own natural
// content height — computed up front via estimateCardHeightEmu so no cell
// has to be rendered twice to find out its size.
function renderLayoutBlock(block, rctx, box, nextShapeId) {
  const { cells, layoutGroup } = block;
  const columns = Math.max(1, layoutGroup.mode === 'grid' ? (layoutGroup.columns ?? cells.length) : cells.length);
  const gapEmu = layoutGroup.gapPt != null ? Math.round(layoutGroup.gapPt * EMU_PER_PT) : BOX_GAP_EMU;
  const colWidth = Math.max(1, Math.floor((box.cx - gapEmu * (columns - 1)) / columns));
  const rowHeight = Math.max(1, Math.min(box.cy, Math.max(...cells.map((cellNodes) => estimateCardHeightEmu(cellNodes, cellNodes[0].padding, colWidth, rctx)))));
  const rowCount = Math.ceil(cells.length / columns);

  let xml = '';
  cells.forEach((cellNodes, idx) => {
    const col = idx % columns;
    const row = Math.floor(idx / columns);
    const cellBox = {
      x: box.x + col * (colWidth + gapEmu),
      y: box.y + row * (rowHeight + gapEmu),
      cx: colWidth,
      cy: rowHeight,
    };
    const result = renderCardShape(cellNodes, rctx, cellBox, nextShapeId(), { forcedCy: rowHeight });
    xml += result.xml;
  });

  return { xml, cy: rowCount * rowHeight + (rowCount - 1) * gapEmu };
}

// A text run's own height has no equivalent of an image's real fitted `cy`
// or a card's own estimateCardHeightEmu — PowerPoint auto-fits/wraps text
// within whatever box it's given rather than this exporter measuring exact
// line breaks itself. estimateCardBlockHeightPt (the same per-block-type
// math a card's own content already uses) is the best available proxy: sum
// each node's own estimated line count at the slide's real content width,
// so 3 short paragraphs advance the cursor a little and a long list
// advances it a lot, rather than every text run claiming a fixed fraction
// of whatever's left regardless of how much text it actually holds.
function estimateTextRunHeightEmu(nodes, maxWidthEmu, rctx) {
  const maxWidthPt = maxWidthEmu / EMU_PER_PT;
  const heightPt = nodes.reduce((sum, n) => sum + estimateCardBlockHeightPt(n, maxWidthPt, rctx) + CARD_BLOCK_SPACING_PT, 0);
  return Math.max(1, Math.round(heightPt * EMU_PER_PT));
}

// Splits a text unit's own `nodes` into however many leading nodes fit
// within `availableCy` EMU, plus whatever's left over — the node-level
// counterpart of classifyContentUnits' own unit-level grouping, used by
// renderSlidePage's text-unit branch when a unit doesn't fit the space
// remaining on the current slide. Walks the same per-node
// estimateCardBlockHeightPt + CARD_BLOCK_SPACING_PT math
// estimateTextRunHeightEmu sums in one shot, just accumulated incrementally
// so a cut point can be found.
//
// `forceFirst` mirrors renderSlidePage's own "a page's first unit always
// renders" tolerance, at node granularity: when true, at least one node is
// always kept even if it alone exceeds `availableCy` — so a page with
// nothing on it yet still makes forward progress instead of an oversized
// single node bouncing forever between "rendered" and "still remaining".
// When false (this text unit isn't the page's first) and even the first
// node alone doesn't fit, `fit` comes back empty so the caller can defer
// the whole unit exactly like a block/media unit that doesn't fit at all.
function splitTextRunNodes(nodes, maxWidthEmu, rctx, availableCy, forceFirst) {
  const maxWidthPt = maxWidthEmu / EMU_PER_PT;
  const fit = [];
  let usedEmu = 0;
  for (let i = 0; i < nodes.length; i++) {
    const nodeEmu = Math.max(1, Math.round((estimateCardBlockHeightPt(nodes[i], maxWidthPt, rctx) + CARD_BLOCK_SPACING_PT) * EMU_PER_PT));
    if (fit.length === 0 && nodeEmu > availableCy && !forceFirst) return { fit: [], rest: nodes };
    if (fit.length > 0 && usedEmu + nodeEmu > availableCy) return { fit, rest: nodes.slice(i) };
    fit.push(nodes[i]);
    usedEmu += nodeEmu;
  }
  return { fit, rest: [] };
}

// Splits one slide GROUP (a run of top-level nodes decided by
// collectSlideGroups — a structural decision, not a "does it fit"
// decision) into its title node plus its already-classified content units.
// Exported so pptx/index.js can classify a group exactly once, then loop
// renderSlidePage over the result across as many actual slides as the
// content needs (see renderSlidePage's own note on why one group no longer
// always means exactly one slide).
export function classifyGroup(group) {
  const { title, rest } = splitTitleAndBody(group);
  return { title, units: classifyContentUnits(rest) };
}

// Renders as many `units` (in original document order) as fit within one
// slide's vertical space, starting right after `title` (if given — only a
// group's first page ever gets one; see pptx/index.js's pagination loop).
// Returns whatever didn't fit as `remainingUnits` so the caller can spin up
// a fresh continuation slide for it. Every group used to map to exactly one
// slide unconditionally, no matter how much content it held — content that
// didn't fit within `geometry.bottomY` just silently rendered past the
// bottom margin instead of flowing onto a next slide the way a real
// multi-page document would (a short intro paragraph plus a tall table
// plus a closing line could all end up on one visually overflowing slide).
//
// A page always renders at least its first unit even when that unit alone
// overflows (`cursor > 0` below) — otherwise an oversized single unit would
// never fit on ANY page and bounce forever between "rendered" and "still
// remaining"; it's left to overflow that one page's own bottom margin
// instead, the same "not a full layout engine" acceptance already used
// throughout this exporter (e.g. box.cy's own clamp just below).
export function renderSlidePage(title, units, rctx, geometry) {
  const b = new XmlBuilder();
  b.el('p:sld', { 'xmlns:a': NS_A, 'xmlns:p': NS_P, 'xmlns:r': NS_R });
  b.el('p:cSld');
  b.el('p:spTree');
  b.el('p:nvGrpSpPr');
  b.self('p:cNvPr', { id: 1, name: '' });
  b.self('p:cNvGrpSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:grpSpPr').up();

  let shapeId = 2;
  let contentTop = geometry.titleY;

  if (title) {
    const titleBox = { x: geometry.margin, y: geometry.titleY, cx: geometry.contentWidth, cy: geometry.titleHeight };
    b.raw(shapeXml('p:sp', { shapeId: shapeId++, name: 'Title', phType: 'title', box: titleBox }, renderBodyBlockParagraphs(title, rctx).join('')));
    contentTop = geometry.bodyY;
  }

  // Walks `units` in their original document order, sharing ONE real
  // vertical cursor across text/blocks/media alike — each unit gets
  // whatever space is actually left on the slide (not a pre-divided
  // fraction), and the cursor advances by however tall it actually
  // rendered (an image's own real fitted size, a card's own estimated
  // content height, or a text run's own estimated line count), so a
  // heading immediately followed by its own image renders as a heading
  // immediately followed by its own reasonably-sized image, matching DOCX's
  // own flowed order instead of clumping all text together, then all media
  // afterward in equal slices regardless of size (the previous behavior).
  let bodyPlaceholderClaimed = false;
  let cursor = 0;
  for (; cursor < units.length; cursor++) {
    const unit = units[cursor];
    const box = { x: geometry.margin, y: contentTop, cx: geometry.contentWidth, cy: Math.max(1, geometry.bottomY - contentTop) };
    let result;
    let claimsPlaceholder = false;
    if (unit.kind === 'block') {
      const block = unit.block;
      if (block.kind === 'layout') {
        result = renderLayoutBlock(block, rctx, box, () => shapeId++);
      } else if (block.kind === 'card') {
        const cardShapeId = shapeId++;
        // A second shape id is only spent when there's actually a borderLeft
        // to draw as an accent bar (see renderAccentBarXml) — not allocated
        // unconditionally, so a plain card without one doesn't burn/skip an id.
        const accentShapeId = block.nodes[0].borderLeft ? shapeId++ : undefined;
        result = renderCardShape(block.nodes, rctx, box, cardShapeId, { accentShapeId });
      } else {
        result = renderStyledBoxShape(block.node, rctx, box, shapeId++);
      }
    } else if (unit.kind === 'media') {
      result = renderMediaShape(unit.node, rctx, box, () => shapeId++);
    } else {
      // Only the FIRST text run on a slide can claim the real body
      // placeholder (a slide layout defines exactly one, idx 1) — every
      // text run after that (e.g. "paragraph, image, paragraph, image")
      // renders as a plain freestanding text box instead (see shapeXml's
      // own note on the phType-less case).
      //
      // A text unit can span several paragraphs/headings/lists merged by
      // classifyContentUnits' own pushText (e.g. every <p> in a scrollable
      // HTML container that had no images/cards interleaved) — when the
      // whole unit doesn't fit what's left on this page, split it at a
      // node boundary via splitTextRunNodes instead of rendering it whole
      // and unclamped (the previous behavior — see that function's own
      // comment for why an unclamped, unclipped `p:sp` was the only way to
      // avoid PowerPoint silently shrinking real text to fit a pre-clamped
      // box). Only the fitting leading nodes render here; the rest becomes
      // this same slot's replacement unit (`units[cursor]`) so the existing
      // `units.slice(cursor)` remainder logic below picks it up for a
      // continuation slide, without this unit ever being silently dropped.
      const availableCy = geometry.bottomY - contentTop;
      const { fit: nodesToRender, rest: remainderNodes } = splitTextRunNodes(unit.nodes, geometry.contentWidth, rctx, availableCy, cursor === 0);
      if (nodesToRender.length === 0) {
        // Nothing fits and this isn't the page's first unit — defer the
        // whole unit, exactly as the pre-split behavior already did.
        break;
      }
      const cy = estimateTextRunHeightEmu(nodesToRender, geometry.contentWidth, rctx);
      const textBox = { ...box, cy };
      const paragraphs = nodesToRender.flatMap((n) => renderBodyBlockParagraphs(n, rctx)).join('');
      claimsPlaceholder = !bodyPlaceholderClaimed;
      const xml = shapeXml('p:sp', {
        shapeId: shapeId++, name: 'Content',
        phType: claimsPlaceholder ? 'body' : undefined, phIdx: claimsPlaceholder ? 1 : undefined,
        box: textBox,
      }, paragraphs);
      result = { xml, cy };
      if (remainderNodes.length > 0) {
        units[cursor] = { kind: 'text', nodes: remainderNodes };
        b.raw(result.xml);
        if (claimsPlaceholder) bodyPlaceholderClaimed = true;
        contentTop += result.cy + BOX_GAP_EMU;
        break; // the remainder, now at units[cursor], continues on a fresh slide
      }
    }
    // This unit doesn't fit what's left on THIS page — defer it (and every
    // unit after it, since document order must stay intact) to a
    // continuation slide instead. Skipped for the page's very first unit
    // (cursor === 0) — see this function's own doc comment on why. A `text`
    // unit never reaches this check when it needed splitting — it already
    // rendered and broke out above — only block/media units (genuinely
    // unsplittable) and whole-fitting text units reach it.
    if (cursor > 0 && contentTop + result.cy > geometry.bottomY) break;
    if (result.xml) b.raw(result.xml);
    if (claimsPlaceholder) bodyPlaceholderClaimed = true;
    contentTop += result.cy + BOX_GAP_EMU;
  }

  b.up(); // p:spTree
  b.up(); // p:cSld
  b.el('p:clrMapOvr');
  b.self('a:overrideClrMapping');
  b.up();
  return { xml: xmlDoc(b.end()), remainingUnits: units.slice(cursor) };
}
