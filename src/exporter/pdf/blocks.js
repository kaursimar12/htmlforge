/**
 * IDM node → PDF drawing calls. Scoped to the PDF1-PDF10 roadmap milestones
 * (headings, paragraphs, inline formatting, lists, page breaks, images,
 * tables, hyperlinks, metadata) plus a first "box styling" pass for
 * paragraph/heading nodes — backgroundColor, border (uniform + per-side),
 * shadow, rounded/circular corners, margin, padding, cardGroup (several
 * sibling boxes sharing one drawn panel), and shrinkToFit (auto-width
 * boxes, e.g. buttons) — the same fields DOCX/PPTX already read (see their
 * own shared/shape-style.js and dom-to-idm.js's mapElement). Real CSS
 * grid/flex column layout, background gradients, anchorId bookmarks, and
 * box styling on table/list/quote/code/section nodes themselves are a
 * separate, larger follow-up, not attempted in this pass.
 */
import { estimateTextWidth, estimateWrappedLines } from '../shared/font-metrics.js';
import { expandTableGrid } from '../shared/table-grid.js';
import { tableRows, flattenCellText } from '../../idm/flatten-text.js';
import { pickBaseFont } from './fonts.js';
import {
  HEADING_SIZES_PT,
  LIST_INDENT_STEP_PT,
  LIST_MARKER_GAP_PT,
  DEFAULT_TABLE_BORDER_RGB,
  DEFAULT_TABLE_CELL_PADDING_PT,
  BLOCK_SPACING_PT,
  SHADOW_OFFSET_PT,
  SHADOW_RGB,
  uaDefaultMarginPt,
} from './layout.js';

const LINK_COLOR_RGB = [0, 0, 0.933];
const CODE_BG_RGB = [0.11, 0.11, 0.13];
const CODE_TEXT_RGB = [0.93, 0.93, 0.93];
const QUOTE_BG_RGB = [0.93, 0.95, 1];
const QUOTE_BORDER_RGB = [0.15, 0.31, 0.85];
const TABLE_HEADER_BG_RGB = [0.9, 0.9, 0.92];

// A `<code>` run always DRAWS in Courier regardless of its own/ambient
// font-family (see fonts.js's pickBaseFont — "code always wins") — every
// WIDTH ESTIMATE for that same run must agree, or the estimated width these
// callers use to position the next segment / decide a wrap point drifts
// from the real drawn width. Underestimating (measuring a Courier run
// against a narrower proportional-font table) is the dangerous direction:
// the next segment starts too early and visibly overlaps the tail of the
// code run — exactly what a mixed "plain text <code>tag</code> more text"
// line does today without this. Distinct from pickBaseFont's own
// fontFamily resolution: this only ever feeds estimateTextWidth/
// estimateWrappedLines, never the actual drawing.
// Falls all the way through to 'Helvetica' (never undefined) when neither
// the run nor the caller declares a family — pickBaseFont's own fallback
// for anything not actually embedded is ALWAYS Helvetica (never the generic,
// character-class-bucketed estimate elsewhere in this module), so the WIDTH
// ESTIMATE for that same undeclared-family run has real per-glyph Adobe
// Core-14 metrics to consult too, instead of a cruder approximation that
// doesn't know it's specifically Helvetica being drawn.
function measurementFontFamily(style, fallback) {
  if (style?.code) return 'Courier New';
  return style?.fontFamily ?? fallback ?? 'Helvetica';
}

function hexToRgb01(hex) {
  if (!hex) return undefined;
  const clean = hex.replace('#', '');
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  if ([r, g, b].some(Number.isNaN)) return undefined;
  return [r, g, b];
}

// ─── Inline run flattening ──────────────────────────────────────────────
// Walks inline IDM children accumulating {bold,italic,code,color,fontFamily,
// href} down through nesting — the same accumulate-then-flatten pattern
// DOCX's renderRuns uses — into one ordered list of leaf {text, style}
// runs. `lineBreak` becomes its own marker entry so callers can hard-split
// on it before wrapping.
function flattenRuns(nodes, style) {
  const runs = [];
  for (const node of nodes ?? []) {
    switch (node.type) {
      case 'text':
        if (node.value) runs.push({ text: node.value, style });
        break;
      case 'strong':
        runs.push(...flattenRuns(node.children, { ...style, bold: true }));
        break;
      case 'emphasis':
        runs.push(...flattenRuns(node.children, { ...style, italic: true }));
        break;
      case 'codeSpan':
        if (node.value) runs.push({ text: node.value, style: { ...style, code: true } });
        break;
      case 'link':
        // An in-document `#anchor` href has no named-destination navigation
        // in this pass (documented non-goal, matching PPTX's own hyperlink
        // renderer) — rendered as fully plain text, no annotation, no
        // link-blue color, so it's indistinguishable from surrounding prose.
        if (node.href?.startsWith('#')) {
          runs.push(...flattenRuns(node.children, style));
        } else {
          runs.push(...flattenRuns(node.children, { ...style, href: node.href, color: LINK_COLOR_RGB }));
        }
        break;
      case 'styledSpan': {
        const s = node.style ?? {};
        runs.push(...flattenRuns(node.children, {
          ...style,
          bold: style.bold || s.bold,
          italic: style.italic || s.italic,
          color: s.color ? hexToRgb01(s.color) : style.color,
          fontFamily: s.fontFamily ?? style.fontFamily,
          // The numeric weight (e.g. 700), distinct from the `bold` boolean
          // above — picking the correct file out of a multi-weight embedded
          // font family (e.g. Poppins-Medium/500 vs Poppins-Bold/700) needs
          // the real number; ">=600 → bold" can't distinguish a 600 from an
          // 800. See inline-style.js's extractRunStyle for where this is set.
          fontWeight: s.fontWeight ?? style.fontWeight,
          // Previously dropped entirely here — <u>/<mark>/<sub>/<sup>/<s> all
          // parse correctly into a styledSpan's own style (see
          // dom-to-idm.js's mapInlineElementCore), but never made it past
          // this accumulator, so every drawn run silently lost its
          // underline/highlight/strikethrough/sub-or-superscript regardless
          // of what the source HTML declared.
          underline: style.underline || s.underline,
          strikethrough: style.strikethrough || s.strikethrough,
          backgroundColor: s.backgroundColor ? hexToRgb01(s.backgroundColor) : style.backgroundColor,
          verticalAlign: s.verticalAlign ?? style.verticalAlign,
          // An explicit CSS font-size (`<span style="font-size:20px">`, or a
          // <p>/<td>'s own style — both wrap their content in exactly this
          // same styledSpan, see dom-to-idm.js's mapInlineWithStyle) was
          // silently dropped here — segmentSizeAndOffset (below) already
          // reads `seg.style.fontSizePt` to override the block's own base
          // size, but a run never carried this far enough to reach it, so
          // every explicit font-size rendered at the surrounding block's
          // plain size instead (flow.bodyFontPt for a paragraph, the
          // heading-level scale for a heading) regardless of what the
          // source HTML declared.
          fontSizePt: s.fontSizePt ?? style.fontSizePt,
        }));
        break;
      }
      case 'lineBreak':
        runs.push({ lineBreak: true });
        break;
      // A cell holding BLOCK content instead of plain inline runs (most
      // commonly several `<p>`s, or an `<h3>` label followed by `<p>`s —
      // e.g. this project's own "From"/"Bill To" address block, or the
      // invoice header's own `<h1>INVOICE</h1><p>...</p><p>...</p>`) used to
      // fall straight through to the `default: break` below, silently
      // dropping the cell's ENTIRE text — the same block-content gap XLSX's
      // own flattenCellRuns and PPTX's own realTableCellXml already needed
      // fixing for. Recursing into each block's own children, with a hard
      // line break inserted between siblings (mirroring how a browser's own
      // block-level layout visually separates them), is that same fix here.
      // A heading/paragraph/quote/section carrying its OWN block-level
      // backgroundColor (e.g. this project's own `<h3 style="background:
      // #1F4E79;color:white">From</h3>` inside a "From"/"Bill To" cell —
      // mapElement stamps the background directly on the node, separate
      // from the paired `color:white` a nested styledSpan already carries
      // via that case's own backgroundColor/color forwarding just above)
      // needs THIS node's own background forwarded too, or the paired white
      // text renders with nothing to sit on — invisible white-on-white
      // against the page's own background, not just plain/unstyled. Unlike
      // XLSX (a whole-cell-fill-only format, where the same pairing has to
      // be dropped entirely — see rich-text.js's dropBackgroundPairedColor)
      // or PPTX's own tight run-highlight approximation, PDF can draw a
      // real background rectangle at any position, so drawTableRow's own
      // per-line drawing (below) promotes this into an actual full-cell-
      // width colored bar behind the whole line, not just a swatch tight
      // to the glyphs.
      case 'heading':
        if (runs.length) runs.push({ lineBreak: true });
        // Same per-level size scale a TOP-LEVEL heading already renders at
        // (see renderHeading's own identical HEADING_SIZES_PT lookup) — a
        // heading flattened into a cell's inline runs still reads as an
        // actual heading, not just bold body text, once segmentSizeAndOffset
        // (below) and wrapSection's own per-line heightPt (see cellLines/
        // drawTableRow) pick this size up.
        runs.push(...flattenRuns(node.children, {
          ...style,
          bold: true,
          fontSizePt: HEADING_SIZES_PT[node.level] ?? HEADING_SIZES_PT[6],
          backgroundColor: node.backgroundColor ? hexToRgb01(node.backgroundColor) : style.backgroundColor,
        }));
        break;
      case 'paragraph':
      case 'quote':
      case 'section':
        if (runs.length) runs.push({ lineBreak: true });
        runs.push(...flattenRuns(node.children, {
          ...style,
          backgroundColor: node.backgroundColor ? hexToRgb01(node.backgroundColor) : style.backgroundColor,
        }));
        break;
      case 'list':
        (node.items ?? []).forEach((item, i) => {
          if (runs.length) runs.push({ lineBreak: true });
          runs.push({ text: node.ordered ? `${i + (node.start ?? 1)}. ` : '• ', style });
          runs.push(...flattenRuns(item.children, style));
        });
        break;
      default:
        break;
    }
  }
  return runs;
}

function splitOnHardBreaks(runs) {
  const sections = [[]];
  for (const run of runs) {
    if (run.lineBreak) sections.push([]);
    else sections[sections.length - 1].push(run);
  }
  return sections;
}

// Splits runs into whitespace-delimited words while preserving which style
// each character came from — e.g. `<em>italic</em>,` is ONE word (["italic"
// (em), ","(plain)]) since there's no actual whitespace between the two
// runs, but `<em>italic</em> ,` (a real space in the source) is two. This is
// what lets wrapSection glue a style change to directly-adjacent text with
// no visible gap, instead of treating "different run" as if it always meant
// "different word" (see wrapSection's own note on why that distinction
// matters — the bug this replaces silently inserted a phantom space at
// every style boundary, e.g. turning `H<sub>2</sub>O` into "H 2 O").
function tokenizeRunsIntoWords(runs) {
  const words = [];
  let current = null;
  for (const run of runs) {
    for (const token of run.text.match(/\S+|\s+/g) ?? []) {
      if (/^\s/.test(token)) {
        if (current) { words.push(current); current = null; }
      } else {
        if (!current) current = [];
        current.push({ text: token, style: run.style });
      }
    }
  }
  if (current) words.push(current);
  return words;
}

/**
 * Wraps one section's runs to `maxWidthPt`, returning `[{segments: [{text,
 * style, gapBeforePt}], maxSizePt, heightPt}]` — one entry per drawn line.
 * `estimateWrappedLines` decides line boundaries using `measureSizePt` — the
 * largest `fontSizePt` any run in this section actually declares (e.g. a
 * heading run flattenRuns sized up via HEADING_SIZES_PT), falling back to
 * the caller's own flat `sizePt` when no run overrides it — using the flat
 * `sizePt` here regardless (the previous behavior) under-estimated how much
 * width a genuinely bigger-font run needs, letting more text than actually
 * fits get crammed onto one line. Since wrapping never reorders words, each
 * computed line's own word count tells us exactly how many *words* (not
 * runs — see tokenizeRunsIntoWords) to pull off the front of a flat,
 * style-tagged word queue built from the same runs. `gapBeforePt` is 0 for a
 * style-run glued to the previous one within the same word (no space to
 * draw) and one space-width for the start of a genuinely new word —
 * callers add it before drawing each segment instead of assuming a uniform
 * space between every one. Each returned line's own `maxSizePt` (the
 * largest actually-drawn size among its own segments — see
 * segmentSizeAndOffset) and `heightPt` (that size scaled by `lineHeightRatio`,
 * the same size-to-leading ratio the caller's own flat lineHeightPt/sizePt
 * already uses) let a caller that mixes line sizes within one flow (e.g.
 * drawTableRow's own table cells — a heading line followed by plain body
 * text) reserve/advance by each line's REAL height instead of a flat one
 * that's only correct when every line happens to share the same font size.
 */
function wrapSection(runs, maxWidthPt, sizePt, baseFontFamily, lineHeightRatio = 1) {
  const measureSizePt = Math.max(sizePt, ...runs.map((r) => r.style?.fontSizePt ?? 0));
  const words = tokenizeRunsIntoWords(runs);
  const fullText = words.map((pieces) => pieces.map((p) => p.text).join('')).join(' ');
  // This wrap decision measures the section's flattened text as ONE string,
  // with no per-run style boundary left to consult by the time it reaches
  // estimateWrappedLines — so, same reasoning as cellMeasurementFontFamily's
  // own "any codeSpan forces Courier for the whole cell", ANY bold run
  // anywhere in the section is enough to measure the WHOLE section as bold.
  // Bold is real always-wider glyph data (see HELVETICA_BOLD_WIDTHS), so
  // this is the safe overestimate direction: a mixed bold/plain line wraps
  // very slightly earlier than it strictly needs to, rather than the
  // dangerous direction (measuring bold text as if it were narrower/plain)
  // that let a bold run's real drawn width run past where wrapping/the next
  // segment assumed it would end.
  const anyBold = runs.some((r) => r.style?.bold);
  const lines = estimateWrappedLines(fullText, maxWidthPt, measureSizePt, baseFontFamily, anyBold);
  const spaceWidth = estimateTextWidth(' ', measureSizePt, baseFontFamily, anyBold);

  let cursor = 0;
  return lines.map((line) => {
    const wordCount = line.length ? line.split(' ').length : 0;
    const lineWords = words.slice(cursor, cursor + wordCount);
    cursor += wordCount;

    const segments = [];
    lineWords.forEach((pieces, wi) => {
      pieces.forEach((piece, pi) => {
        const isWordStart = pi === 0;
        const last = segments[segments.length - 1];
        if (last && sameStyle(last.style, piece.style)) {
          // Same style as whatever's already queued immediately before —
          // merge into one drawn segment (fewer draw calls). A genuinely new
          // word (isWordStart && not the line's first word) needs a real
          // space character folded into the merged text itself; a
          // continuation piece within the same word needs none at all.
          last.text += (isWordStart && wi > 0 ? ' ' : '') + piece.text;
        } else {
          const gapBeforePt = segments.length === 0 ? 0 : isWordStart ? spaceWidth : 0;
          segments.push({ text: piece.text, style: piece.style, gapBeforePt });
        }
      });
    });
    const maxSizePt = Math.max(sizePt, ...segments.map((seg) => segmentSizeAndOffset(seg, sizePt).sizePt));
    return { segments, maxSizePt, heightPt: maxSizePt * lineHeightRatio };
  });
}

function sameStyle(a, b) {
  return a.bold === b.bold && a.italic === b.italic && a.code === b.code && a.href === b.href
    && a.color?.join() === b.color?.join() && a.fontFamily === b.fontFamily
    && a.underline === b.underline && a.strikethrough === b.strikethrough && a.verticalAlign === b.verticalAlign
    && a.backgroundColor?.join() === b.backgroundColor?.join();
}

// Sub/superscript render at a reduced size with the baseline nudged down/up
// — the standard typographic convention CSS's own UA default approximates
// (roughly 0.7em size, with a -0.15em/+0.25em baseline shift) — computed
// per-segment rather than for the whole block, so `H<sub>2</sub>O` mid-line
// only shrinks the "2", not the surrounding text. `seg.style.fontSizePt`
// (e.g. a heading run flattenRuns sized up via HEADING_SIZES_PT) overrides
// the caller's own flat `sizePt` as this segment's OWN base size before
// sub/superscript scaling applies — without this override, a heading
// flattened into a table cell's inline runs (see flattenRuns's own
// 'heading' case) rendered at the surrounding body text's size, merely
// bold, with no size distinction to show it was ever a heading at all.
function segmentSizeAndOffset(seg, sizePt) {
  const base = seg.style.fontSizePt ?? sizePt;
  if (seg.style.verticalAlign === 'subscript') return { sizePt: base * 0.7, yOffset: -base * 0.15 };
  if (seg.style.verticalAlign === 'superscript') return { sizePt: base * 0.7, yOffset: base * 0.25 };
  return { sizePt: base, yOffset: 0 };
}

/**
 * Draws one already-positioned inline segment — its own text plus whatever
 * decorations its style carries: a highlight background (drawn first, so
 * the text sits on top of it), subscript/superscript sizing (see
 * segmentSizeAndOffset), an underline/strikethrough rule, and a hyperlink's
 * own underline + clickable annotation. Shared by every per-segment draw
 * loop in this file (drawWrappedLines, renderList, renderQuote, table
 * cells) so all of this behaves identically no matter which block type the
 * segment came from — previously each loop drew only plain colored text,
 * silently dropping underline/highlight/sub/sup a styledSpan's own style
 * already carried (see flattenRuns's own now-expanded accumulation).
 * Returns the segment's own drawn width, in points, for the caller to
 * advance its cursor by.
 */
function drawSegment(flow, x, baselineY, seg, { sizePt, baseFontFamily, forceBold = false, forceItalic = false, defaultColorRGB } = {}) {
  const { sizePt: segSizePt, yOffset } = segmentSizeAndOffset(seg, sizePt);
  const y = baselineY + yOffset;
  const fontFamily = seg.style.fontFamily ?? baseFontFamily;
  // Computed before the width estimate below (not after, as this used to
  // read) — a forced bold (e.g. a table header cell's own forceBold) must
  // widen the WIDTH ESTIMATE too, or the next segment's start position is
  // computed as if this one drew at its narrower regular-weight width while
  // it's actually drawn via pickBaseFont's own bold variant right below,
  // visually colliding with whatever comes right after it.
  const bold = forceBold || seg.style.bold;
  const italic = forceItalic || seg.style.italic;
  const segWidth = estimateTextWidth(seg.text, segSizePt, measurementFontFamily(seg.style, baseFontFamily), bold);
  const colorRGB = seg.style.color ?? defaultColorRGB;

  if (seg.style.backgroundColor) {
    flow.drawRect(x, y - segSizePt * 0.25, segWidth, segSizePt * 1.15, { fillRGB: seg.style.backgroundColor });
  }

  // A `<code>` run always draws Courier regardless of the surrounding
  // font-family (matches this project's existing DOCX/PPTX "code is always
  // monospace" convention — see fonts.js's own pickBaseFont) — so an
  // embedded-font lookup is skipped entirely for code runs, same as if no
  // fontFamily had ever been declared.
  const fontResolution = !seg.style.code && fontFamily
    ? flow.resolveTextFont(fontFamily, { weight: seg.style.fontWeight ?? (bold ? 700 : 400), italic })
    : undefined;
  const embedded = fontResolution?.kind === 'embedded' ? flow.useEmbeddedFont(fontResolution.entry) : undefined;

  if (embedded) {
    flow.drawEmbeddedText(x, y, seg.text, { resourceName: embedded.resourceName, descriptor: embedded.descriptor, sizePt: segSizePt, colorRGB });
  } else {
    const baseFont = pickBaseFont({ bold, italic, code: seg.style.code });
    flow.drawText(x, y, seg.text, { baseFont, sizePt: segSizePt, colorRGB });
  }

  if (seg.style.href) {
    flow.drawLine(x, y - 1, x + segWidth, y - 1, { strokeRGB: LINK_COLOR_RGB, lineWidth: 0.5 });
    flow.addLinkAnnotation([x, y - sizePt * 0.2, x + segWidth, y + sizePt * 0.8], seg.style.href);
  } else if (seg.style.underline) {
    flow.drawLine(x, y - 1, x + segWidth, y - 1, { strokeRGB: colorRGB ?? [0, 0, 0], lineWidth: 0.5 });
  }
  if (seg.style.strikethrough) {
    flow.drawLine(x, y + segSizePt * 0.3, x + segWidth, y + segSizePt * 0.3, { strokeRGB: colorRGB ?? [0, 0, 0], lineWidth: 0.5 });
  }

  return segWidth;
}

/**
 * Draws one already-wrapped paragraph's worth of lines, one `ensureSpace`
 * call per line — the finest natural unit text already loops over, which is
 * what makes "a paragraph taller than a page" a non-event (see page-flow.js).
 */
function drawWrappedLines(flow, sectionsOfLines, { sizePt, lineHeightPt, align, defaultColorRGB, baseBold, baseItalic, baseFontFamily, left = flow.marginLeft, width = flow.contentWidth }) {
  for (const lines of sectionsOfLines) {
    for (const line of lines) {
      flow.ensureSpace(lineHeightPt);
      const baselineY = flow.cursorY - sizePt * 0.85; // ~ascender offset from the line's own top

      const segmentWidths = line.segments.map((seg) => estimateTextWidth(seg.text, segmentSizeAndOffset(seg, sizePt).sizePt, measurementFontFamily(seg.style, baseFontFamily), baseBold || seg.style.bold));
      const totalWidth = segmentWidths.reduce((a, w) => a + w, 0) + line.segments.reduce((a, seg) => a + seg.gapBeforePt, 0);

      let x = left;
      if (align === 'center') x = left + Math.max(0, (width - totalWidth) / 2);
      else if (align === 'right') x = left + Math.max(0, width - totalWidth);
      // 'justify' is treated as 'left' — real inter-word stretch justification
      // is out of scope for this pass, same "not a full layout engine"
      // simplification the rest of this project already accepts elsewhere.

      for (const seg of line.segments) {
        x += seg.gapBeforePt;
        x += drawSegment(flow, x, baselineY, seg, { sizePt, baseFontFamily, forceBold: baseBold, forceItalic: baseItalic, defaultColorRGB });
      }

      flow.advance(lineHeightPt);
    }
  }
}

function renderRunsAsBlock(flow, children, { sizePt, bold = false, italic = false, align, colorRGB, fontFamily }) {
  const runs = flattenRuns(children, { bold, italic });
  const sections = splitOnHardBreaks(runs).map((section) => wrapSection(section, flow.contentWidth, sizePt, fontFamily));
  const lineHeightPt = sizePt * flow.lineHeightMultiplier;
  drawWrappedLines(flow, sections, { sizePt, lineHeightPt, align, defaultColorRGB: colorRGB, baseBold: bold, baseItalic: italic, baseFontFamily: fontFamily });
}

// ─── Box styling (backgroundColor/border/shadow/rounded/padding/margin) ──
// The fields dom-to-idm.js's mapElement bleeds onto a paragraph/heading node
// when the source `<div>`/etc. it collapsed from carried CSS box styling
// (see that function's own extensive notes) — the same fields DOCX's
// resolveEdgeBorders/renderPBdr/renderSpacing and PPTX's shared
// shape-style.js already read, reproduced here as real drawn PDF geometry
// instead of OOXML/DrawingML markup.

function computeWrappedSections(node, sizePt, availableWidth, bold) {
  const runs = flattenRuns(node.children, { bold });
  return splitOnHardBreaks(runs).map((section) => wrapSection(section, availableWidth, sizePt, undefined));
}

function countLines(sectionsOfLines) {
  return sectionsOfLines.reduce((sum, lines) => sum + lines.length, 0) || 1;
}

function maxLineWidthPt(sectionsOfLines, sizePt) {
  let max = 0;
  for (const lines of sectionsOfLines) {
    for (const line of lines) {
      const total = line.segments.reduce((sum, seg) => sum + estimateTextWidth(seg.text, segmentSizeAndOffset(seg, sizePt).sizePt, measurementFontFamily(seg.style), seg.style.bold) + seg.gapBeforePt, 0);
      if (total > max) max = total;
    }
  }
  return max;
}

// A uniform `border` always wins over per-side borders for the same reason
// extractBlockSideBorder's own doc comment gives: a uniform border is a
// stronger, more specific "this is one bounded panel" signal, and the two
// are mutually exclusive in how the parser populates them (a plain `border`
// shorthand never also produces `borderTop`/etc.).
function nodeBoxStyle(node) {
  const backgroundColor = hexToRgb01(node.backgroundColor);
  const border = node.border ? { color: hexToRgb01(node.border.color), widthPt: node.border.sizeEighthPt / 8 } : undefined;
  const sideBorders = !border && (node.borderTop || node.borderRight || node.borderBottom || node.borderLeft)
    ? { top: node.borderTop, right: node.borderRight, bottom: node.borderBottom, left: node.borderLeft }
    : undefined;
  const padding = node.padding ?? {};
  const hasBox = Boolean(backgroundColor || border || sideBorders || node.shadow);
  return {
    backgroundColor, border, sideBorders, hasBox,
    shadow: node.shadow, rounded: node.rounded, circular: node.circular,
    padding: { top: padding.top ?? 0, right: padding.right ?? 0, bottom: padding.bottom ?? 0, left: padding.left ?? 0 },
  };
}

// `x`/`y` is the box's bottom-left corner (PDF's own rect convention — see
// drawRect), `w`/`h` its full size. Shadow draws first (offset down-right,
// behind everything), then the fill+uniform-border pass, then each declared
// side's own thin accent-bar rect on top — matching PPTX's own
// renderAccentBarXml precedent for a lone `border-left` stripe, generalized
// to whichever sides are actually set.
function drawNodeBox(flow, x, y, w, h, box) {
  if (box.shadow) {
    flow.drawRect(x + SHADOW_OFFSET_PT, y - SHADOW_OFFSET_PT, w, h, { fillRGB: SHADOW_RGB, rounded: box.rounded, circular: box.circular });
  }
  if (box.backgroundColor || box.border) {
    flow.drawRect(x, y, w, h, {
      fillRGB: box.backgroundColor,
      strokeRGB: box.border?.color,
      lineWidth: box.border?.widthPt,
      rounded: box.rounded,
      circular: box.circular,
    });
  }
  if (box.sideBorders) {
    const s = box.sideBorders;
    if (s.top) flow.drawRect(x, y + h - s.top.sizeEighthPt / 8, w, s.top.sizeEighthPt / 8, { fillRGB: hexToRgb01(s.top.color) });
    if (s.bottom) flow.drawRect(x, y, w, s.bottom.sizeEighthPt / 8, { fillRGB: hexToRgb01(s.bottom.color) });
    if (s.left) flow.drawRect(x, y, s.left.sizeEighthPt / 8, h, { fillRGB: hexToRgb01(s.left.color) });
    if (s.right) flow.drawRect(x + w - s.right.sizeEighthPt / 8, y, s.right.sizeEighthPt / 8, h, { fillRGB: hexToRgb01(s.right.color) });
  }
}

/**
 * Computes everything renderStyledRunBlock needs to draw one paragraph/
 * heading — sections, box style, sizes — as a pure function of `node` and
 * `availableWidth`, with NO drawing. Shared by the real renderer and
 * measureBlockSequenceHeight's grid/flex pre-measurement pass, so a grid
 * cell's required row height is computed with the exact same math the real
 * draw uses, not a separate approximation that could disagree with it.
 */
function measureRunBlock(node, sizePt, bold, availableWidth, lineHeightMultiplier) {
  const box = nodeBoxStyle(node);
  const ua = uaDefaultMarginPt(node.type, node.level);
  const beforePt = node.margin?.beforePt ?? ua.beforePt;
  const afterPt = node.margin?.afterPt ?? ua.afterPt;
  const lineHeightPt = sizePt * lineHeightMultiplier;

  if (!box.hasBox) {
    const sections = computeWrappedSections(node, sizePt, availableWidth, bold);
    return { box, beforePt, afterPt, sections, lineHeightPt, contentHeight: countLines(sections) * lineHeightPt };
  }

  const pad = box.padding;
  const innerWidth = Math.max(1, availableWidth - pad.left - pad.right);
  const sections = computeWrappedSections(node, sizePt, innerWidth, bold);
  const textHeight = countLines(sections) * lineHeightPt + pad.top + pad.bottom;
  // An explicit CSS `height` on the boxed element (e.g. a full-height
  // `.hero{height:900px}` banner) is a real authored size — never shrinking
  // the box below what its own text actually needs, so real wrapped text is
  // never clipped by an under-estimated explicit height. `extraPt` is the
  // slack between the two, advanced past (see renderStyledRunBlock) after
  // the text so the flow cursor still lands exactly at the drawn box's own
  // bottom edge, not partway up it.
  const explicitHeightPt = node.heightPx != null ? node.heightPx * PX_TO_PT : undefined;
  const contentHeight = Math.max(textHeight, explicitHeightPt ?? 0);
  const extraPt = contentHeight - textHeight;
  return { box, beforePt, afterPt, sections, lineHeightPt, contentHeight, pad, innerWidth, extraPt };
}

/**
 * Renders one paragraph/heading, drawing its own background/border/shadow
 * box first (inset by padding) when `node` carries any box styling, else
 * falling back to the plain unboxed text flow every non-styled block already
 * used. `node.margin.{beforePt,afterPt}` replaces the flat BLOCK_SPACING_PT
 * gap when present, matching DOCX's renderSpacing/PPTX's box-aware spacing.
 */
function renderStyledRunBlock(flow, node, sizePt, bold) {
  const m = measureRunBlock(node, sizePt, bold, flow.contentWidth, flow.lineHeightMultiplier);
  flow.applySpacingBefore(m.beforePt);

  if (!m.box.hasBox) {
    drawWrappedLines(flow, m.sections, { sizePt, lineHeightPt: m.lineHeightPt, align: node.align, baseBold: bold });
    flow.deferSpacingAfter(m.afterPt);
    return;
  }

  // `shrinkToFit` (a button, or an explicit `display:inline-block`/fixed
  // width) shrinks the drawn box to its own content instead of stretching
  // across the full content width — the same signal DOCX's
  // renderShrinkToFitBox/PPTX's renderStyledBoxShape key off of.
  let boxWidth = flow.contentWidth;
  let boxX = flow.marginLeft;
  // `m.innerWidth` (measureRunBlock's own wrapping width) is based on the
  // FULL available width — correct for the common stretched-box case, but
  // once shrinkToFit narrows boxWidth below that, drawWrappedLines' own
  // center/right-align math (`left + (width - textWidth)`) must use THAT
  // narrower width too, or it offsets the text by the full, unshrunk width
  // on top of boxX already having shifted right/center to the shrunk box's
  // own position — compounding into text landing far past the shrunk box's
  // own right edge, in the worst case (a page-wide `shrinkToFit` bled from
  // an unrelated ancestor's own explicit `width`, paired with a right-
  // aligned line) pushing it entirely off the page and rendering nothing
  // visible at all. Left-aligned shrinkToFit text never exposed this: its
  // own x is always just `left` regardless of `width`.
  let innerWidth = m.innerWidth;
  if (node.shrinkToFit) {
    boxWidth = Math.min(flow.contentWidth, maxLineWidthPt(m.sections, sizePt) + m.pad.left + m.pad.right);
    innerWidth = Math.max(1, boxWidth - m.pad.left - m.pad.right);
    if (node.align === 'center') boxX = flow.marginLeft + (flow.contentWidth - boxWidth) / 2;
    else if (node.align === 'right') boxX = flow.marginLeft + flow.contentWidth - boxWidth;
  }

  // A styled box taller than one full page is the same accepted
  // pathological case as an oversized table row (see renderTable) — draw
  // what fits on this page rather than trying to split a background/border
  // across a page boundary.
  const pageContentHeight = flow.geometry.contentTop - flow.geometry.contentBottom;
  flow.ensureSpace(Math.min(m.contentHeight, pageContentHeight));

  const top = flow.cursorY;
  drawNodeBox(flow, boxX, top - m.contentHeight, boxWidth, m.contentHeight, m.box);

  flow.advance(m.pad.top);
  drawWrappedLines(flow, m.sections, { sizePt, lineHeightPt: m.lineHeightPt, align: node.align, baseBold: bold, left: boxX + m.pad.left, width: innerWidth });
  flow.advance(m.pad.bottom);
  if (m.extraPt) flow.advance(m.extraPt);
  flow.deferSpacingAfter(m.afterPt);
}

function renderHeading(flow, node) {
  const sizePt = HEADING_SIZES_PT[node.level] ?? HEADING_SIZES_PT[6];
  renderStyledRunBlock(flow, node, sizePt, true);
}

function renderParagraph(flow, node) {
  renderStyledRunBlock(flow, node, flow.bodyFontPt, false);
}

// The same pure-measurement/render split as measureRunBlock/renderStyledRunBlock above.
function measureCardGroup(members, availableWidth, bodyFontPt, lineHeightMultiplier) {
  const box = nodeBoxStyle(members[0]);
  // Deliberately NOT falling back to uaDefaultMarginPt for the OUTER
  // beforePt/afterPt here, unlike every other block type below: a card
  // group draws a REAL box around its members (see drawNodeBox), and CSS
  // margin collapsing rules say a child's own margin does NOT collapse
  // through a parent that has real padding — it stays INSIDE that padding
  // instead of becoming space outside the parent's box. Modeled below as
  // real spacing between padding and the first/last member and between
  // members themselves (gapBeforePt/trailingGapPt) — see there for why that
  // internal placement is exactly where a member's own margin belongs.
  const beforePt = members[0].margin?.beforePt ?? 0;
  const afterPt = members[members.length - 1].margin?.afterPt ?? 0;
  const pad = box.padding;
  const innerWidth = Math.max(1, availableWidth - pad.left - pad.right);

  const runBlocks = [];
  for (const node of members) {
    if (node.type !== 'paragraph' && node.type !== 'heading') continue;
    const sizePt = node.type === 'heading' ? (HEADING_SIZES_PT[node.level] ?? HEADING_SIZES_PT[6]) : bodyFontPt;
    const bold = node.type === 'heading';
    const sections = computeWrappedSections(node, sizePt, innerWidth, bold);
    const lineHeightPt = sizePt * lineHeightMultiplier;
    const ua = uaDefaultMarginPt(node.type, node.level);
    runBlocks.push({
      node, sizePt, bold, sections, lineHeightPt,
      height: countLines(sections) * lineHeightPt,
      memberBeforePt: node.margin?.beforePt ?? ua.beforePt,
      memberAfterPt: node.margin?.afterPt ?? ua.afterPt,
    });
  }

  // A member's own margin (explicit CSS, e.g. a `<p style="margin-top:...">`
  // inside the card, or the browser's own UA-default heading/paragraph
  // margin) stays INSIDE the box — see the note above — landing as real
  // space between the padding and the first member, between consecutive
  // members (collapsed to the max of the two, same sibling-collapsing rule
  // used everywhere else in this renderer), and between the last member and
  // the padding. This is what makes a hero-style `<div><h1>Title</h1>
  // <p>Subtitle</p></div>` card's own real height match a browser's —
  // without it, h1/p drew back-to-back with zero gap between them, and the
  // browser's own default ~1em heading margin never showed up anywhere.
  runBlocks.forEach((b, i) => {
    b.gapBeforePt = i === 0 ? b.memberBeforePt : Math.max(runBlocks[i - 1].memberAfterPt, b.memberBeforePt);
  });
  const trailingGapPt = runBlocks.length ? runBlocks[runBlocks.length - 1].memberAfterPt : 0;

  const contentHeight = runBlocks.reduce((sum, b) => sum + b.gapBeforePt + b.height, 0) + trailingGapPt;
  const naturalHeight = contentHeight + pad.top + pad.bottom;
  // An explicit CSS `height` on the source `<div>` (e.g. a full-height
  // `.hero{height:900px}` banner wrapping a heading plus a paragraph) is a
  // real authored size — dom-to-idm.js's mapElement bleeds it onto every
  // member of the group, so any member (here `members[0]`) carries it. Never
  // shrinks below what the group's own content actually needs. `extraPt` is
  // the slack between the two, advanced past (see renderCardGroup) after the
  // content so the flow cursor still lands exactly at the drawn box's own
  // bottom edge, not partway up it.
  const explicitHeightPt = members[0].heightPx != null ? members[0].heightPx * PX_TO_PT : undefined;
  const boxHeight = Math.max(naturalHeight, explicitHeightPt ?? 0);
  const extraPt = boxHeight - naturalHeight;
  return { box, beforePt, afterPt, pad, innerWidth, runBlocks, trailingGapPt, boxHeight, extraPt };
}

/**
 * Several sibling paragraph/heading nodes that all collapsed from the same
 * source `<div>` container (dom-to-idm.js's `cardGroupId` — a shared object
 * identity token, compared by reference, not value) share ONE drawn
 * background/border/shadow box instead of each getting its own — the PDF
 * analogue of DOCX's renderCardShapeDrawing / PPTX's grouped renderCardShape
 * for a heading-plus-paragraphs "card".
 */
function renderCardGroup(flow, members, imageXObjectByAssetId) {
  const m = measureCardGroup(members, flow.contentWidth, flow.bodyFontPt, flow.lineHeightMultiplier);
  flow.applySpacingBefore(m.beforePt);

  const pageContentHeight = flow.geometry.contentTop - flow.geometry.contentBottom;
  flow.ensureSpace(Math.min(m.boxHeight, pageContentHeight));

  const top = flow.cursorY;
  drawNodeBox(flow, flow.marginLeft, top - m.boxHeight, flow.contentWidth, m.boxHeight, m.box);

  flow.advance(m.pad.top);
  for (const b of m.runBlocks) {
    flow.advance(b.gapBeforePt);
    drawWrappedLines(flow, b.sections, {
      sizePt: b.sizePt, lineHeightPt: b.lineHeightPt, align: b.node.align, baseBold: b.bold,
      left: flow.marginLeft + m.pad.left, width: m.innerWidth,
    });
  }
  flow.advance(m.trailingGapPt);
  flow.advance(m.pad.bottom);
  if (m.extraPt) flow.advance(m.extraPt);
  flow.deferSpacingAfter(m.afterPt);

  // A cardGroup member that isn't itself a paragraph/heading (rare — the
  // parser only ever stamps cardGroup on BOX_STYLE_NODE_TYPES, which
  // includes table/list/quote/code/figure too) still needs to render, just
  // without sharing the box drawn above.
  for (const node of members) {
    if (node.type !== 'paragraph' && node.type !== 'heading') renderBlock(flow, node, imageXObjectByAssetId);
  }
}

// ─── CSS grid/flex layout (layoutGroup/layoutItemIndex) ──────────────────
// dom-to-idm.js's mapLayoutBlocks stamps every paragraph/heading a
// `display:grid`/`flex` container's own direct children produced with a
// shared `layoutGroup` object (compared by reference, like cardGroup) and a
// `layoutItemIndex` — the position of the *source child* that node came
// from, so a child that itself expanded to several nodes (e.g. a `.box` div
// with an `<h3>` plus a bare text line) still counts as exactly one cell.
// Real layout engines aside, this project's "no unequal track sizing, no
// wrapping" simplification is the same one PPTX's own renderLayoutBlock and
// DOCX's renderLayoutTable already accept — every row gets N equal-width
// columns, `layoutGroup.columns` many per row (or all cells in one row for a
// plain `display:flex` with no declared track count).

function measureBlockSequenceHeight(nodes, availableWidth, bodyFontPt, lineHeightMultiplier) {
  let total = 0;
  let i = 0;
  while (i < nodes.length) {
    const node = nodes[i];
    if (node.cardGroup) {
      let j = i + 1;
      while (j < nodes.length && nodes[j].cardGroup === node.cardGroup) j++;
      const m = measureCardGroup(nodes.slice(i, j), availableWidth, bodyFontPt, lineHeightMultiplier);
      total += m.beforePt + m.boxHeight + m.afterPt;
      i = j;
    } else if (node.type === 'paragraph' || node.type === 'heading') {
      const sizePt = node.type === 'heading' ? (HEADING_SIZES_PT[node.level] ?? HEADING_SIZES_PT[6]) : bodyFontPt;
      const m = measureRunBlock(node, sizePt, node.type === 'heading', availableWidth, lineHeightMultiplier);
      total += m.beforePt + m.contentHeight + m.afterPt;
      i++;
    } else {
      // A list/table/image/quote/code inside a grid/flex cell is rare and
      // not cheaply pre-measurable without duplicating its own layout math
      // — approximate with a fixed block height; ensureSpace still protects
      // the real draw from actually overflowing a page.
      total += bodyFontPt * lineHeightMultiplier * 3 + BLOCK_SPACING_PT;
      i++;
    }
  }
  return total;
}

/**
 * Renders one row of `columns` side-by-side cells (each cell itself a
 * node sequence, rendered via the ordinary renderBlockSequence so a cell's
 * own cardGroup/box styling still applies) — every cell starts at the same
 * `rowTopY` inside its own `withHorizontalSlot`, and the shared cursor
 * advances by whichever cell ended up tallest once all are drawn.
 */
function renderLayoutRow(flow, cells, columns, gapPt, imageXObjectByAssetId) {
  const totalGap = gapPt * (columns - 1);
  const columnWidth = Math.max(1, (flow.contentWidth - totalGap) / columns);

  // Flushes whatever the previous block (or previous row, for a grid with
  // more items than columns) deferred, before measuring/drawing this row.
  flow.applySpacingBefore(0);

  const rowHeight = Math.max(0, ...cells.map((cellNodes) => measureBlockSequenceHeight(cellNodes, columnWidth, flow.bodyFontPt, flow.lineHeightMultiplier)));
  const pageContentHeight = flow.geometry.contentTop - flow.geometry.contentBottom;
  flow.ensureSpace(Math.min(rowHeight, pageContentHeight));

  const rowTopY = flow.cursorY;
  const outerLeft = flow.marginLeft;
  let lowestY = rowTopY;
  cells.forEach((cellNodes, c) => {
    const cellX = outerLeft + c * (columnWidth + gapPt);
    flow.cursorY = rowTopY;
    // Each column is an independent vertical flow starting fresh at the
    // row's own top — without this, column 2's first block would collapse
    // its own margin against whatever column 1 happened to leave pending,
    // even though they're side by side, not stacked.
    flow.resetPendingSpacing();
    flow.withHorizontalSlot(cellX, columnWidth, () => {
      renderBlockSequence(flow, cellNodes, imageXObjectByAssetId);
    });
    if (flow.cursorY < lowestY) lowestY = flow.cursorY;
  });
  flow.resetPendingSpacing(); // discard whatever the last cell deferred — the row's own trailing gap (below) is independent
  flow.cursorY = lowestY;
  // Reuses the container's own CSS `gap` value for the row-to-row spacing
  // too — real CSS `gap` sets row-gap and column-gap to the same value
  // unless overridden, so this is still HTML-declared spacing, not an
  // invented default.
  flow.deferSpacingAfter(gapPt);
}

// A `display:grid`/`flex` run's cells, grouped by layoutItemIndex (several
// nodes from the same source child share one cell), then chunked into rows
// of `layoutGroup.columns` (or all cells in one row when the container
// never declared `grid-template-columns` — the plain `display:flex` case).
function renderLayoutGroup(flow, groupNodes, layoutGroup, imageXObjectByAssetId) {
  const cellsByIndex = new Map();
  for (const node of groupNodes) {
    const idx = node.layoutItemIndex ?? 0;
    // Consumed here — each cell's own node sequence is rendered via
    // renderBlockSequence below (renderLayoutRow), which would otherwise see
    // this same `layoutGroup` reference still set on every member and
    // re-enter renderLayoutGroup with the exact same nodes forever.
    node.layoutGroup = undefined;
    if (!cellsByIndex.has(idx)) cellsByIndex.set(idx, []);
    cellsByIndex.get(idx).push(node);
  }
  const cells = [...cellsByIndex.keys()].sort((a, b) => a - b).map((idx) => cellsByIndex.get(idx));

  // Every node sharing one `layoutItemIndex` came from the SAME source child
  // of the grid/flex container (mapLayoutBlocks increments the index once
  // per child, not per node that child expanded to) — a stronger "these
  // belong together" signal than dom-to-idm's own cross-sibling `cardGroup`
  // heuristic even requires. That heuristic's own `trustedPanel` gate (needs
  // a real `border`, or `backgroundColor` paired with an explicit `color`)
  // deliberately doesn't trust a plain `.box{background:#eef4ff}` with
  // neither — so a grid cell's own `<h3>`+text pair (each independently
  // picking up the SAME backgroundColor from their shared `.box` parent,
  // since that bleed has no such gate) would otherwise draw as two
  // visually disconnected boxes stacked instead of one card. Synthesizing a
  // cardGroup here — gated on at least one member actually having box
  // styling, so a plain unstyled multi-node cell keeps its normal
  // between-block spacing instead of being tightened for no visual reason —
  // merges them into the single shared box the source markup clearly meant.
  for (const cellNodes of cells) {
    const eligible = cellNodes.filter((n) => n.type === 'paragraph' || n.type === 'heading');
    if (eligible.length > 1 && !eligible.some((n) => n.cardGroup) && eligible.some((n) => nodeBoxStyle(n).hasBox)) {
      const syntheticGroup = {};
      for (const n of eligible) n.cardGroup = syntheticGroup;
    }
  }

  const columns = layoutGroup.columns > 0 ? layoutGroup.columns : cells.length;
  const gapPt = layoutGroup.gapPt ?? 12;
  for (let start = 0; start < cells.length; start += columns) {
    renderLayoutRow(flow, cells.slice(start, start + columns), columns, gapPt, imageXObjectByAssetId);
  }
}

/**
 * Walks a sequence of sibling block nodes (document.children, or a
 * `section`'s own children), grouping consecutive same-`layoutGroup` members
 * into a real side-by-side grid/flex row (renderLayoutGroup) and consecutive
 * same-`cardGroup` members into one shared `renderCardGroup` box — both
 * compared by reference (see nodeBoxStyle's own note) — before falling back
 * to rendering a plain node independently.
 */
function renderBlockSequence(flow, nodes, imageXObjectByAssetId) {
  let i = 0;
  while (i < nodes.length) {
    const node = nodes[i];
    if (node.layoutGroup) {
      let j = i + 1;
      while (j < nodes.length && nodes[j].layoutGroup === node.layoutGroup) j++;
      renderLayoutGroup(flow, nodes.slice(i, j), node.layoutGroup, imageXObjectByAssetId);
      i = j;
    } else if (node.cardGroup) {
      let j = i + 1;
      while (j < nodes.length && nodes[j].cardGroup === node.cardGroup) j++;
      renderCardGroup(flow, nodes.slice(i, j), imageXObjectByAssetId);
      i = j;
    } else {
      renderBlock(flow, node, imageXObjectByAssetId);
      i++;
    }
  }
}

// Real browsers apply NO background/border/italic of their own to a plain
// `<blockquote>` — only the UA stylesheet's own `margin:1em 40px` indent
// (see uaDefaultMarginPt's own note on where that pattern comes from). This
// is the browser-default LEFT/RIGHT half of that same rule (the vertical
// 1em half is uaDefaultMarginPt('quote')'s own beforePt/afterPt) — used only
// when the source declares no padding/indent of its own, same "explicit
// beats inferred" precedence as everywhere else in this renderer.
const UA_QUOTE_INDENT_PT = 30;

function renderQuote(flow, node, imageXObjectByAssetId) {
  const uaQuote = uaDefaultMarginPt('quote');
  flow.applySpacingBefore(node.margin?.beforePt ?? uaQuote.beforePt);
  const sizePt = flow.bodyFontPt;
  const lineHeightPt = sizePt * flow.lineHeightMultiplier;
  const box = nodeBoxStyle(node);
  // A real declared border (uniform, or left-only — the .note/.warn accent
  // pattern) is the gate for "this blockquote is genuinely meant to be a
  // highlighted box," NOT box.hasBox/backgroundColor alone: a plain
  // background bleeding down from an unrelated page-level wrapper (see
  // dom-to-idm.js's mapElement — background bleeds onto every unclaimed
  // child as a plain "highlight" even with no border, deliberately) would
  // otherwise make an entirely unstyled `<blockquote>` sitting inside such a
  // wrapper draw a border/accent bar it never asked for. Same "border is the
  // strong, specific panel signal; background alone isn't" reasoning
  // dom-to-idm.js's own trustedPanel/cardGroupId already use.
  const declaredBorder = box.border ?? (box.sideBorders?.left ? { color: hexToRgb01(box.sideBorders.left.color), widthPt: box.sideBorders.left.sizeEighthPt / 8 } : undefined);
  const isStyled = Boolean(declaredBorder);
  // A blockquote whose source CSS declares no border of its own (the common
  // case — see the note above) gets NO invented "highlighted pull-quote"
  // box: plain, non-italic, own-color text at the browser's default indent,
  // nothing drawn behind it. Only a genuinely bordered blockquote gets any
  // box at all, and then using ITS OWN declared colors, not a hardcoded
  // look forced onto every quote regardless of what the HTML asked for.
  const p = node.padding ?? {};
  const leftIndent = isStyled ? (p.left ?? LIST_INDENT_STEP_PT) : (p.left ?? UA_QUOTE_INDENT_PT);
  const rightPad = isStyled ? (p.right ?? 6) : (p.right ?? UA_QUOTE_INDENT_PT);
  const topPad = p.top ?? 0;
  const bottomPad = p.bottom ?? 0;
  const accentColor = declaredBorder?.color ?? QUOTE_BORDER_RGB;
  const accentWidthPt = declaredBorder?.widthPt ?? 3;
  for (const child of node.children ?? []) {
    if (child.type !== 'paragraph') { renderBlock(flow, child, imageXObjectByAssetId); continue; }
    const runs = flattenRuns(child.children, {});
    const availableWidth = flow.contentWidth - leftIndent - rightPad;
    const lines = splitOnHardBreaks(runs).flatMap((section) => wrapSection(section, availableWidth, sizePt, undefined));
    const boxHeight = lines.length * lineHeightPt + topPad + bottomPad;

    flow.ensureSpace(Math.min(boxHeight, flow.geometry.contentTop - flow.geometry.contentBottom));
    const top = flow.cursorY;
    if (isStyled) {
      flow.drawRect(flow.marginLeft, top - boxHeight, flow.contentWidth, boxHeight, { fillRGB: box.backgroundColor ?? QUOTE_BG_RGB });
      flow.drawRect(flow.marginLeft, top - boxHeight, accentWidthPt, boxHeight, { fillRGB: accentColor });
    }

    let lineY = top - topPad - sizePt * 0.85;
    for (const line of lines) {
      let x = flow.marginLeft + leftIndent;
      for (const seg of line.segments) {
        x += seg.gapBeforePt;
        x += drawSegment(flow, x, lineY, seg, { sizePt });
      }
      lineY -= lineHeightPt;
    }
    flow.advance(boxHeight);
  }
  flow.deferSpacingAfter(node.margin?.afterPt ?? uaQuote.afterPt);
}

// Code blocks render one drawn line per *source* line, unwrapped — matching
// this project's existing convention elsewhere (DOCX joins source lines
// with a literal w:br, PPTX draws one paragraph per source line); wrapping
// a code line would misrepresent the source text.
function renderCode(flow, node) {
  const uaCode = uaDefaultMarginPt('code');
  flow.applySpacingBefore(node.margin?.beforePt ?? uaCode.beforePt);
  const sizePt = flow.bodyFontPt;
  const lineHeightPt = sizePt * flow.lineHeightMultiplier;
  const bg = hexToRgb01(node.backgroundColor) ?? CODE_BG_RGB;
  const color = hexToRgb01(node.color) ?? CODE_TEXT_RGB;
  // Falls back to this renderer's own small fixed inset only when the
  // source `<pre>` declared no padding of its own — see
  // codeFromStyledElement's own note on why that's now read directly off
  // the element instead of being silently dropped.
  const p = node.padding ?? {};
  const pad = { top: p.top ?? 6, right: p.right ?? 6, bottom: p.bottom ?? 6, left: p.left ?? 6 };
  const lines = node.value.split('\n');
  const boxHeight = lines.length * lineHeightPt + pad.top + pad.bottom;

  // One shared rect for the whole block (not one per line, now that real
  // top/bottom padding insets it) — a code block taller than a page is the
  // same accepted "draw what fits, don't split a background across a page
  // boundary" case renderStyledRunBlock's own styled-box path already uses.
  const pageContentHeight = flow.geometry.contentTop - flow.geometry.contentBottom;
  flow.ensureSpace(Math.min(boxHeight, pageContentHeight));
  const top = flow.cursorY;
  flow.drawRect(flow.marginLeft, top - boxHeight, flow.contentWidth, boxHeight, { fillRGB: bg });

  let lineY = top - pad.top - sizePt * 0.85;
  for (const sourceLine of lines) {
    flow.drawText(flow.marginLeft + pad.left, lineY, sourceLine, { baseFont: pickBaseFont({ code: true }), sizePt, colorRGB: color });
    lineY -= lineHeightPt;
  }
  flow.advance(boxHeight);
  flow.deferSpacingAfter(node.margin?.afterPt ?? uaCode.afterPt);
}

function listMarker(list, index) {
  if (!list.ordered) return '•';
  return `${(list.start ?? 1) + index}.`;
}

function renderList(flow, list, depth, imageXObjectByAssetId) {
  const uaList = uaDefaultMarginPt('list');
  if (depth === 0) flow.applySpacingBefore(list.margin?.beforePt ?? uaList.beforePt);
  const sizePt = flow.bodyFontPt;
  const lineHeightPt = sizePt * flow.lineHeightMultiplier;
  const indent = LIST_INDENT_STEP_PT * (depth + 1);

  list.items.forEach((item, index) => {
    const [first, ...rest] = item.children ?? [];
    if (first?.type === 'paragraph') {
      const marker = listMarker(list, index);
      const markerWidth = estimateTextWidth(marker, sizePt, undefined);
      const availableWidth = flow.contentWidth - indent - markerWidth - LIST_MARKER_GAP_PT;
      const runs = flattenRuns(first.children, {});
      const lines = splitOnHardBreaks(runs).flatMap((section) => wrapSection(section, availableWidth, sizePt, undefined));

      lines.forEach((line, lineIndex) => {
        flow.ensureSpace(lineHeightPt);
        const baselineY = flow.cursorY - sizePt * 0.85;
        if (lineIndex === 0) {
          flow.drawText(flow.marginLeft + indent - markerWidth - LIST_MARKER_GAP_PT, baselineY, marker, { baseFont: pickBaseFont({}), sizePt });
        }
        let x = flow.marginLeft + indent;
        for (const seg of line.segments) {
          x += seg.gapBeforePt;
          x += drawSegment(flow, x, baselineY, seg, { sizePt });
        }
        flow.advance(lineHeightPt);
      });
    } else if (first) {
      renderBlock(flow, first, imageXObjectByAssetId);
    }

    for (const child of rest) {
      if (child.type === 'list') renderList(flow, child, depth + 1, imageXObjectByAssetId);
      else renderBlock(flow, child, imageXObjectByAssetId);
    }
  });
  flow.deferSpacingAfter(list.margin?.afterPt ?? uaList.afterPt);
}

function renderThematicBreak(flow) {
  flow.applySpacingBefore(0);
  const gap = 6;
  flow.ensureSpace(gap * 2);
  const y = flow.cursorY - gap;
  flow.drawLine(flow.marginLeft, y, flow.marginLeft + flow.contentWidth, y, { strokeRGB: [0.6, 0.6, 0.6], lineWidth: 1 });
  flow.advance(gap * 2);
}

// ─── Tables ─────────────────────────────────────────────────────────────

// A cell's own explicit CSS padding (`th,td{padding:8px}`, already extracted
// onto `cell.padding` by dom-to-idm.js's mapRows) always wins — the flat
// DEFAULT_TABLE_CELL_PADDING_PT is only the fallback for a side the cell
// never declared, matching the CSS-first precedence this project now uses
// everywhere else (margins, box padding, etc.), not a PDF-specific override
// of whatever the HTML actually says.
function cellPaddingPt(cell) {
  const p = cell?.padding;
  return {
    top: p?.top ?? DEFAULT_TABLE_CELL_PADDING_PT,
    right: p?.right ?? DEFAULT_TABLE_CELL_PADDING_PT,
    bottom: p?.bottom ?? DEFAULT_TABLE_CELL_PADDING_PT,
    left: p?.left ?? DEFAULT_TABLE_CELL_PADDING_PT,
  };
}

function cellLines(cell, colWidthPt, sizePt, lineHeightRatio) {
  if (!cell) return [];
  // `cell.header`'s forced bold (see drawTableRow's own forceBold: cell?.header)
  // has to be seeded into the base style here too — otherwise the WRAP
  // decision measures a header cell's text as regular-weight while it's
  // actually drawn bold, under-counting how many lines it needs exactly like
  // an explicit `font-weight:bold` cell would without this.
  const runs = flattenRuns(cell.children, { bold: cell.bold || cell.header, italic: cell.italic });
  const pad = cellPaddingPt(cell);
  const availableWidth = Math.max(1, colWidthPt - pad.left - pad.right);
  const baseFontFamily = cellMeasurementFontFamily(cell);
  return splitOnHardBreaks(runs).flatMap((section) => wrapSection(section, availableWidth, sizePt, baseFontFamily, lineHeightRatio));
}

// A cell's own `<img>` (bare, or a `<figure>`'s single image — its caption is
// dropped in this pass, same scope limit as everywhere else images-in-cells
// weren't handled at all) never turns into a flattenRuns text run — an image
// XObject needs real x/y/w/h placement, not a text glyph position. Nested
// tables are skipped: an image inside a NESTED table belongs to that inner
// table's own cell layout (drawn when that inner table itself renders as a
// block), not stacked into this outer cell's own image area.
function collectCellImageNodes(nodes) {
  const out = [];
  const walk = (list) => {
    for (const node of list ?? []) {
      if (!node) continue;
      if (node.type === 'image') { out.push(node); continue; }
      if (node.type === 'figure') { if (node.image) out.push(node.image); continue; }
      if (node.type === 'table') continue;
      if (node.children) walk(node.children);
      if (node.items) for (const item of node.items) walk(item.children);
    }
  };
  walk(nodes);
  return out;
}

const CELL_IMAGE_GAP_PT = 6; // gap between a cell's own text block and its image stack, and between stacked images

// Mirrors renderImage's own "real intrinsic size always wins over a
// declared width/height" precedence (see that function's own doc comment)
// and aspect-fit scaling, scaled down to the cell's own available width
// instead of the page's full content width — a `<td>` is almost always
// narrower than the page, so scaling against the page width would routinely
// overflow the cell.
function cellImageBoxes(cell, availableWidthPt, imageXObjectByAssetId) {
  if (!cell) return [];
  const boxes = [];
  for (const node of collectCellImageNodes(cell.children)) {
    const xObject = imageXObjectByAssetId.get(node.assetId);
    if (!xObject) continue; // unresolved or unsupported format — already warned
    // xObject only ever exists for a fully-decoded JPEG/PNG (buildImageXObject
    // returns undefined otherwise), so dict.Width/Height are always the
    // image's own real pixel dimensions here — never a guess to fall back
    // from, unlike DOCX/PPTX/XLSX's shared image-size.js which also has to
    // cover an undecodable format with no intrinsic size at all.
    const w = xObject.dict.Width * PX_TO_PT;
    const h = xObject.dict.Height * PX_TO_PT;
    const scale = Math.min(1, availableWidthPt / w);
    boxes.push({ assetId: node.assetId, xObject, w: w * scale, h: h * scale });
  }
  return boxes;
}

function cellImageStackHeightPt(boxes) {
  if (boxes.length === 0) return 0;
  return boxes.reduce((sum, b) => sum + b.h, 0) + (boxes.length - 1) * CELL_IMAGE_GAP_PT;
}

// `cellLines` (via wrapSection/estimateWrappedLines) always returns AT
// LEAST one line even for a cell with zero actual text — an explicit,
// tested contract of estimateWrappedLines itself (empty text still wraps
// to one empty line, so callers never have to special-case "zero lines").
// That means `rowLines[i].length > 0` is ALWAYS true, even for a pure-image
// cell with no text runs at all — checking array length instead of real
// content silently treated every image-only cell as "has a preceding text
// line", permanently disabling the vertical-centering path below and
// mis-sizing the row height by one phantom empty line's worth of space.
// This checks for a line with real drawable segments instead.
function cellHasVisibleText(lines) {
  return lines.some((line) => line.segments.length > 0);
}

// Mirrors DOCX's document.js `CELL_STYLE_FIELDS`/`cellHasExplicitStyle` —
// presence of ANY of these on a cell means the source HTML's author
// deliberately styled this table, just not necessarily with a visible
// border (see tableHasExplicitStyling below for why that distinction
// matters). `align` is handled separately since a `<th>`'s own
// browser-default center-alignment (applied even with zero CSS at all —
// see mapCells' own `isHeader ? 'center' : undefined`) would otherwise
// wrongly count as "the author styled this." Deliberately excludes
// `widthPercent`, unlike DOCX's own list: a bare `<td width="52%">` is a
// pure old-school LAYOUT hint with no bearing on appearance at all — the
// "content pasted with minimal markup, give it a readable default grid"
// case this whole check exists to detect still very much applies to a
// table whose only markup is a column-width hint.
const CELL_STYLE_FIELDS = [
  'backgroundColor', 'borderTop', 'borderRight', 'borderBottom', 'borderLeft', 'noBorder',
  'padding', 'verticalAlign', 'color', 'bold', 'italic', 'underline', 'strikethrough',
  'fontSizePt', 'direction', 'writingMode', 'rotateDeg',
];

function cellHasExplicitStyle(cell) {
  if (!cell) return false;
  if (CELL_STYLE_FIELDS.some((field) => cell[field] !== undefined)) return true;
  return cell.align !== undefined && !(cell.header && cell.align === 'center');
}

// Whether `table` (or any of its cells) carries ANY sign the source HTML
// author styled it deliberately — mirrors DOCX's own `tableHasExplicitStyling`
// identically, including its reasoning: a table with literally zero CSS
// anywhere (the common "content pasted with no styling at all" case) gets
// the flat default gray grid every cell drew unconditionally before this —
// still the sensible, readable default for that case. A table the author DID
// style in some way (this project's own `<table style="width:50%;
// margin-left:auto">` totals table, with per-cell `text-align:right` but
// deliberately no border) should have that specific choice — no border,
// except wherever a cell/table declares one — respected instead of
// smothered under a border it never asked for.
function tableHasExplicitStyling(table, grid) {
  if (table.border != null || table.widthPercent != null || table.marginAuto != null) return true;
  for (const row of grid) {
    for (const slot of row) {
      if (!slot.continuation && cellHasExplicitStyle(slot.cell)) return true;
    }
  }
  return false;
}

// One declared side's border, `cell`'s own if present, else the table's own
// uniform `border` (e.g. `<table style="border:1px solid #ccc">`) as every
// cell's shared default for that side — never applied over an explicit
// `noBorder` (a cell that actively cancelled its border via `border:none`).
function resolveCellEdgeBorder(cell, side, tableBorder) {
  const own = cell?.[`border${side[0].toUpperCase()}${side.slice(1)}`];
  if (own) return own;
  return cell?.noBorder ? undefined : tableBorder;
}

// `x`/`y` is the cell's bottom-left corner (matches drawRect/drawNodeBox's
// own convention), `w`/`h` its full size — same per-side thin-accent-bar
// drawing drawNodeBox already uses for a block's own borderTop/etc.
function drawCellBorders(flow, x, y, w, h, cell, tableBorder) {
  const top = resolveCellEdgeBorder(cell, 'top', tableBorder);
  const right = resolveCellEdgeBorder(cell, 'right', tableBorder);
  const bottom = resolveCellEdgeBorder(cell, 'bottom', tableBorder);
  const left = resolveCellEdgeBorder(cell, 'left', tableBorder);
  if (top) flow.drawRect(x, y + h - top.sizeEighthPt / 8, w, top.sizeEighthPt / 8, { fillRGB: hexToRgb01(top.color) });
  if (bottom) flow.drawRect(x, y, w, bottom.sizeEighthPt / 8, { fillRGB: hexToRgb01(bottom.color) });
  if (left) flow.drawRect(x, y, left.sizeEighthPt / 8, h, { fillRGB: hexToRgb01(left.color) });
  if (right) flow.drawRect(x + w - right.sizeEighthPt / 8, y, right.sizeEighthPt / 8, h, { fillRGB: hexToRgb01(right.color) });
}

function drawTableRow(flow, gridRow, colWidths, sizePt, lineHeightPt, borderMode, startX = flow.marginLeft, imageXObjectByAssetId = new Map()) {
  // `colWidths` is indexed by REAL column position (0..columnCount-1);
  // `gridRow` is indexed by SLOT position, which only lines up 1:1 with a
  // real column when every cell in the row has colspan=1 — expandTableGrid
  // (table-grid.js) doesn't insert a placeholder slot for the columns a
  // colspan covers (unlike a rowspan continuation, which does get one), so
  // a `<td colspan="4">` is just ONE slot needing the SUM of 4 real columns'
  // width, not colWidths at its own slot index alone. Using the slot index
  // directly (the previous behavior) silently gave a colspan cell only its
  // first covered column's own width — invisible while every column
  // happened to be a similar width already, but very visible once columns
  // genuinely differ (e.g. this project's own Items table: an 8%-wide "#"
  // column right next to a `<td colspan="4">GST (18%)</td>` label that
  // actually spans 82% of the table — wrapped "GST (18%)" onto two lines
  // because it was only given the "#" column's own narrow width to work with).
  const slotWidths = [];
  let realCol = 0;
  for (const slot of gridRow) {
    slotWidths.push(colWidths.slice(realCol, realCol + slot.colspan).reduce((a, b) => a + b, 0));
    realCol += slot.colspan;
  }
  // The same size-to-leading ratio the table's own flat lineHeightPt/sizePt
  // already encodes, reused so a cell mixing line sizes (a heading line
  // followed by plain body text — see flattenRuns' own 'heading' case) gets
  // each line's OWN proportionally-scaled height instead of every line
  // being forced through the body text's own flat lineHeightPt regardless
  // of how much taller a heading line actually draws.
  const lineHeightRatio = lineHeightPt / sizePt;
  const rowLines = gridRow.map((slot, i) => (slot.continuation ? [] : cellLines(slot.cell, slotWidths[i], sizePt, lineHeightRatio)));
  const rowPads = gridRow.map((slot) => cellPaddingPt(slot.continuation ? slot.originCell : slot.cell));
  // A cell's own `<img>` (e.g. a product thumbnail or logo cell) reserves
  // its own vertical space in the row too — same "measure everything before
  // committing to one rowHeight" reasoning as rowLines, just for the one
  // content kind flattenRuns/cellLines can never represent (see
  // collectCellImageNodes' own doc comment).
  const rowImages = gridRow.map((slot, i) => (slot.continuation ? [] : cellImageBoxes(slot.cell, Math.max(1, slotWidths[i] - rowPads[i].left - rowPads[i].right), imageXObjectByAssetId)));
  // Sums each cell's own lines' REAL heightPt (see wrapSection) rather than
  // `lines.length * lineHeightPt` — the previous flat multiply silently
  // assumed every line in the cell was the same (body-text) size, so a
  // cell with a genuinely taller heading line got a row too short to fit
  // it, crowding/overlapping the line right after it.
  const rowHeight = Math.max(lineHeightPt, ...rowLines.map((lines, i) => {
    const textHeight = cellHasVisibleText(lines) ? lines.reduce((sum, l) => sum + l.heightPt, 0) : 0;
    const imagesHeight = cellImageStackHeightPt(rowImages[i]);
    const gapBetween = textHeight > 0 && imagesHeight > 0 ? CELL_IMAGE_GAP_PT : 0;
    return textHeight + gapBetween + imagesHeight + rowPads[i].top + rowPads[i].bottom;
  }));

  flow.ensureSpace(rowHeight);
  const top = flow.cursorY;
  let x = startX;

  gridRow.forEach((slot, i) => {
    const cellWidth = slotWidths[i];
    const cell = slot.continuation ? slot.originCell : slot.cell;
    const pad = rowPads[i];
    // A cell's own explicit CSS background (`th{background:#2563eb}`) always
    // wins — the flat gray fallback is only for a header cell that declared
    // no background of its own, matching DOCX's own `w:shd` precedence
    // (explicit backgroundColor first, `header` a distinguishing default
    // second, never the other way around).
    const bg = hexToRgb01(cell?.backgroundColor) ?? (cell?.header ? TABLE_HEADER_BG_RGB : undefined);
    if (borderMode.defaultGrid) {
      flow.drawRect(x, top - rowHeight, cellWidth, rowHeight, { fillRGB: bg, strokeRGB: DEFAULT_TABLE_BORDER_RGB, lineWidth: 0.5 });
    } else {
      flow.drawRect(x, top - rowHeight, cellWidth, rowHeight, { fillRGB: bg });
      drawCellBorders(flow, x, top - rowHeight, cellWidth, rowHeight, cell, borderMode.tableBorder);
    }

    if (!slot.continuation) {
      // A cell's own `text-align` (e.g. this project's own totals table:
      // `<td style="text-align:right;">$1250</td>`) positions each line
      // within the cell's own available width instead of always starting
      // at its left edge — matching how a browser actually renders it.
      const cellAvailableWidth = Math.max(1, cellWidth - pad.left - pad.right);
      // Descends from the cell's own top by that FIRST line's own size (not
      // the table's flat sizePt) — a cell starting with a bigger heading
      // line needs its own taller initial drop to its first baseline, same
      // reasoning as the per-line heightPt used for every advance below.
      let lineY = top - pad.top - (rowLines[i][0]?.maxSizePt ?? sizePt) * 0.85;
      for (const line of rowLines[i]) {
        // A block child's own backgroundColor (see flattenRuns' own note on
        // this exact pairing) is carried on every one of its segments — the
        // first segment's own value stands in for "this whole line shares
        // one block-level background" (the common case: a heading's text is
        // one uniform color/background, not a mix). Drawn as a real
        // FULL-CELL-WIDTH bar behind the line (not just tight around the
        // glyphs — see drawSegment's own per-segment highlight, still used
        // for an actual inline highlight like <mark>) so a heading like
        // "From"/"Bill To" reads as the same colored bar the browser shows,
        // not a swatch that only covers its own short text.
        const lineBg = line.segments[0]?.style.backgroundColor;
        if (lineBg) {
          flow.drawRect(x + pad.left, lineY - line.maxSizePt * 0.25, cellAvailableWidth, line.maxSizePt * 1.15, { fillRGB: lineBg });
        }
        const lineWidth = line.segments.reduce((sum, seg) => sum + estimateTextWidth(seg.text, segmentSizeAndOffset(seg, sizePt).sizePt, measurementFontFamily(seg.style), cell?.header || seg.style.bold) + seg.gapBeforePt, 0);
        let lx = x + pad.left;
        if (cell?.align === 'center') lx += Math.max(0, (cellAvailableWidth - lineWidth) / 2);
        else if (cell?.align === 'right') lx += Math.max(0, cellAvailableWidth - lineWidth);
        for (const seg of line.segments) {
          lx += seg.gapBeforePt;
          lx += drawSegment(flow, lx, lineY, seg, { sizePt, forceBold: cell?.header });
        }
        lineY -= line.heightPt;
      }

      const boxes = rowImages[i];
      if (boxes.length > 0) {
        // A cell holding ONLY image(s) (no text) centers its image stack
        // vertically within the cell's own available height — the common
        // "photo preview" column case (this project's own "Preview"
        // column), matching how a browser vertically centers a lone
        // <img> in a table cell rather than pinning it to the top and
        // leaving all the slack space below whenever the ROW ends up
        // taller than this cell's own image needs (a sibling cell's own
        // taller text content, or another image in the same row, is what
        // usually drives that extra height). A cell with text FOLLOWED by
        // image(s) keeps the image(s) flush after the last text line
        // instead — there's no single obvious "center" for mixed content,
        // so this only applies to the pure-image case.
        let imgY;
        if (cellHasVisibleText(rowLines[i])) {
          imgY = lineY - CELL_IMAGE_GAP_PT;
        } else {
          const cellAvailableHeight = rowHeight - pad.top - pad.bottom;
          const slack = Math.max(0, cellAvailableHeight - cellImageStackHeightPt(boxes));
          imgY = top - pad.top - slack / 2;
        }
        for (const box of boxes) {
          let bx = x + pad.left;
          if (cell?.align === 'center') bx += Math.max(0, (cellAvailableWidth - box.w) / 2);
          else if (cell?.align === 'right') bx += Math.max(0, cellAvailableWidth - box.w);
          flow.registerImage(box.assetId, box.xObject);
          flow.drawImage(box.xObject, bx, imgY - box.h, box.w, box.h);
          imgY -= box.h + CELL_IMAGE_GAP_PT;
        }
      }
    }
    x += cellWidth;
  });

  flow.advance(rowHeight);
}

// ─── Column-width estimation ──────────────────────────────────────────────
// No explicit per-column width declared anywhere: approximate a browser's
// content-based auto table layout (which sizes each column to fit its own
// content) instead of a flat equal split, which badly squeezes a
// short-content column ("Growth") next to a wider one in the same row,
// forcing mid-word wraps a browser never would. Mirrors DOCX's own
// computeContentBasedWidths, using real font-metrics text-width estimation
// (estimateTextWidth) rather than DOCX's flat per-character twips constant,
// since that's already available here and more accurate.
const MIN_COL_WIDTH_PT = 24; // ~1/3in — a pure spacer column still gets a small floor rather than collapsing toward zero

// `cellRequiredWidthPt` measures a cell's ENTIRE flattened text as one plain
// string, with no per-run style boundaries left to consult — so unlike
// measurementFontFamily (which corrects a single already-known run/segment),
// this has to find a single representative family for the whole cell
// first. A `codeSpan` ANYWHERE in the cell means Courier's wider, flat
// per-character width applies to at least part of the real rendered text;
// using it for the WHOLE cell's estimate is the safe overestimate direction
// (a column sized slightly wide is a cosmetic gap, a column sized too
// narrow is the actual "text runs past the table" overflow this fixes) —
// same "safe direction to err" precedent as e.g. WIDE_SCRIPT_RANGES/
// autoSizeBiasPercent elsewhere in this project. Otherwise, the first
// explicit `font-family` found (the common real-world case: a `<td
// style="font-family:...">` wraps its ENTIRE content in one styledSpan, see
// dom-to-idm.js's mapInlineWithStyle) stands in for the whole cell — a cell
// that genuinely mixes several distinct families has no single "right"
// answer here, so this doesn't attempt per-run-accurate measurement, only a
// clear improvement over ALWAYS ignoring font-family entirely.
function cellMeasurementFontFamily(cell) {
  let family;
  let hasCode = false;
  const scan = (nodes) => {
    for (const node of nodes ?? []) {
      if (node.type === 'codeSpan' || node.type === 'code') hasCode = true;
      if (node.type === 'styledSpan') {
        if (family === undefined && node.style?.fontFamily) family = node.style.fontFamily;
        scan(node.children);
      }
    }
  };
  scan(cell?.children);
  return hasCode ? 'Courier New' : family;
}

function cellRequiredWidthPt(cell, sizePt) {
  if (!cell) return 0;
  const pad = cellPaddingPt(cell);
  return estimateTextWidth(flattenCellText(cell), sizePt, cellMeasurementFontFamily(cell), cell.bold || cell.header) + pad.left + pad.right;
}

function computeContentBasedColumnWidths(grid, columnCount, availableWidthPt, sizePt) {
  const minWidths = new Array(columnCount).fill(0);
  const spanningCells = [];

  for (const gridRow of grid) {
    let col = 0;
    for (const slot of gridRow) {
      if (!slot.continuation && slot.cell) {
        if (slot.colspan === 1) {
          minWidths[col] = Math.max(minWidths[col], cellRequiredWidthPt(slot.cell, sizePt));
        } else {
          spanningCells.push({ col, colspan: slot.colspan, cell: slot.cell });
        }
      }
      col += slot.colspan;
    }
  }

  for (const { col, colspan, cell } of spanningCells) {
    const required = cellRequiredWidthPt(cell, sizePt);
    const current = minWidths.slice(col, col + colspan).reduce((a, b) => a + b, 0);
    const deficit = required - current;
    if (deficit > 0) {
      const share = deficit / colspan;
      for (let i = 0; i < colspan; i++) minWidths[col + i] += share;
    }
  }

  const floored = minWidths.map((w) => Math.max(MIN_COL_WIDTH_PT, w));
  const totalMin = floored.reduce((a, b) => a + b, 0);

  if (totalMin >= availableWidthPt) {
    // The table's own minimum content doesn't fit at all — scale every
    // column down proportionally so it still sums to exactly the available
    // width, same as DOCX's own fixed-layout requirement.
    const scale = availableWidthPt / totalMin;
    return floored.map((w) => Math.max(1, w * scale));
  }

  // Distribute whatever's left over beyond everyone's minimum
  // proportionally to each column's own minimum, so a column that already
  // needed more room also gets a proportionally bigger share of any
  // leftover space, rather than every column growing by the same flat amount.
  const slack = availableWidthPt - totalMin;
  const widths = floored.map((w) => w + (w / totalMin) * slack);
  const used = widths.reduce((a, b) => a + b, 0);
  widths[widths.length - 1] += availableWidthPt - used; // absorb any rounding leftover
  return widths;
}

// A layout table's first row typically carries each column's real width as
// a percentage (`<td width="52%">`) — when every cell in row 0 declares
// one, distribute the available width by those proportions instead of
// estimating from content, matching DOCX's own computeColumnWidths precedence.
function computeColumnWidths(grid, columnCount, availableWidthPt, sizePt) {
  const firstRow = grid[0] ?? [];
  const allSpecified = firstRow.length > 0 && firstRow.every((slot) => slot.cell?.widthPercent != null);
  const totalPercent = firstRow.reduce((sum, slot) => sum + (slot.cell?.widthPercent ?? 0), 0);

  if (!allSpecified || totalPercent <= 0) {
    return computeContentBasedColumnWidths(grid, columnCount, availableWidthPt, sizePt);
  }

  const widths = firstRow.map((slot) => (slot.cell.widthPercent / totalPercent) * availableWidthPt);
  const used = widths.reduce((a, b) => a + b, 0);
  widths[widths.length - 1] += availableWidthPt - used;
  return widths;
}

function renderTable(flow, table, imageXObjectByAssetId = new Map()) {
  flow.applySpacingBefore(0);
  const { grid, columnCount } = expandTableGrid(tableRows(table));
  if (grid.length === 0) return;
  const sizePt = flow.bodyFontPt;
  const lineHeightPt = sizePt * flow.lineHeightMultiplier;

  // A table's own `width:N%` (e.g. this project's own "totals" summary
  // table: `width:50%;margin-left:auto`) narrows the drawn table instead of
  // always stretching it across the full content width, matching a
  // browser's own <table> sizing. `margin-left:auto` (with no matching
  // margin-right:auto) floats the narrowed table to the right edge;
  // `margin:auto` on both sides centers it — same two cases
  // extractMarginAuto already distinguishes ('left' vs. 'both').
  const tableWidthPt = table.widthPercent != null ? flow.contentWidth * (table.widthPercent / 100) : flow.contentWidth;
  let startX = flow.marginLeft;
  if (table.marginAuto === 'both') startX = flow.marginLeft + (flow.contentWidth - tableWidthPt) / 2;
  else if (table.marginAuto === 'left') startX = flow.marginLeft + (flow.contentWidth - tableWidthPt);

  const colWidths = computeColumnWidths(grid, columnCount, tableWidthPt, sizePt);
  // A genuinely bare, zero-CSS table gets the flat default gray grid every
  // cell always drew before this — a deliberately-styled one (even if that
  // styling chose no border) instead only ever gets whichever border it
  // actually declared, table-level or per-cell (see tableHasExplicitStyling).
  const borderMode = tableHasExplicitStyling(table, grid)
    ? { defaultGrid: false, tableBorder: table.border }
    : { defaultGrid: true, tableBorder: undefined };

  const headerRowCount = table.head?.length ?? 0;
  const headerRows = grid.slice(0, headerRowCount);

  grid.forEach((gridRow, rowIndex) => {
    const pageBefore = flow.pageIndex;
    drawTableRow(flow, gridRow, colWidths, sizePt, lineHeightPt, borderMode, startX, imageXObjectByAssetId);
    // A page break happened mid-table (detected by comparing page index
    // before/after this row drew) — repeat the table's own declared header
    // row(s), never a "first row" guess, at the top of the continuation
    // page before resuming body rows.
    if (flow.pageIndex !== pageBefore && rowIndex >= headerRowCount && headerRows.length > 0) {
      for (const headerRow of headerRows) drawTableRow(flow, headerRow, colWidths, sizePt, lineHeightPt, borderMode, startX, imageXObjectByAssetId);
    }
  });
  flow.deferSpacingAfter(table.margin?.afterPt ?? 0);
}

// ─── Images ─────────────────────────────────────────────────────────────

const PX_TO_PT = 72 / 96;

// `marginBeforePt`/`marginAfterPt` come from the *figure* node's own CSS
// margin when this image is a `<figure>`'s single `<img>` (see the 'figure'
// dispatch case below) — a bare `<img>` has no wrapping box of its own to
// carry one, same as it has no other box styling (BOX_STYLE_NODE_TYPES never
// includes 'image', only 'figure').
function renderImage(flow, imageNode, imageXObjectByAssetId, captionText, marginBeforePt, marginAfterPt) {
  const xObject = imageXObjectByAssetId.get(imageNode.assetId);
  if (!xObject) return; // unresolved or unsupported format — already warned
  flow.applySpacingBefore(marginBeforePt ?? 0);

  // xObject.dict.Width/Height are the image's own real decoded pixel
  // dimensions — buildImageXObject (images.js) only ever succeeds for a
  // fully-parsed JPEG/PNG, so these are always real measurements, never a
  // guess — and always win outright over whatever the source HTML declared
  // (a stale `<img style="width:80px;height:80px">` pointing at a
  // since-swapped full-res photo shouldn't render at the wrong claimed
  // size), same "real beats declared beats placeholder" precedence
  // DOCX/PPTX/XLSX apply via their own shared exporter/shared/image-size.js.
  let w = xObject.dict.Width * PX_TO_PT;
  let h = xObject.dict.Height * PX_TO_PT;

  const availableWidth = flow.contentWidth;
  const availableHeight = flow.remainingHeight() > 40 ? flow.remainingHeight() : flow.geometry.contentTop - flow.geometry.contentBottom;
  const scale = Math.min(1, availableWidth / w, availableHeight / h);
  w *= scale;
  h *= scale;

  flow.registerImage(imageNode.assetId, xObject);
  flow.ensureSpace(h);
  const x = flow.marginLeft + (availableWidth - w) / 2;
  const y = flow.cursorY - h;
  flow.drawImage(xObject, x, y, w, h);
  flow.advance(h);

  if (captionText) {
    renderRunsAsBlock(flow, [{ type: 'text', value: captionText }], { sizePt: flow.bodyFontPt - 2, italic: true, align: 'center', colorRGB: [0.4, 0.4, 0.4] });
    flow.deferSpacingAfter(marginAfterPt ?? 0);
  } else {
    flow.deferSpacingAfter(marginAfterPt ?? 0);
  }
}

// ─── SVG vector shapes ──────────────────────────────────────────────────
// `node.shapes` coordinates are already in a single [0,widthPx]×[0,heightPx]
// px box (any viewBox already resolved by the parser — see dom-to-idm.js's
// extractSvgShapes) with SVG's own Y-down origin at the top-left. PDF's own
// coordinate space is Y-up from the page's bottom-left, so every point gets
// flipped (not just scaled) when mapped from `originTopY` — the same
// "reserve a block, aspect-fit scale it down, draw within it" pattern
// renderImage already uses for raster images, just with real vector paint
// ops instead of a `Do` XObject call.

function svgToPdfPoint(x, y, originX, originTopY, scale) {
  return [originX + x * scale, originTopY - y * scale];
}

function drawSvgShape(flow, shape, originX, originTopY, scale) {
  const fillRGB = hexToRgb01(shape.fill);
  const strokeRGB = hexToRgb01(shape.stroke);
  const lineWidth = shape.strokeWidthPx != null ? shape.strokeWidthPx * scale : undefined;

  switch (shape.kind) {
    case 'rect': {
      const [x, yBottom] = svgToPdfPoint(shape.x, shape.y + shape.height, originX, originTopY, scale);
      flow.drawRect(x, yBottom, shape.width * scale, shape.height * scale, { fillRGB, strokeRGB, lineWidth });
      break;
    }
    case 'circle': {
      const [x, yBottom] = svgToPdfPoint(shape.cx - shape.r, shape.cy + shape.r, originX, originTopY, scale);
      flow.drawRect(x, yBottom, shape.r * 2 * scale, shape.r * 2 * scale, { fillRGB, strokeRGB, lineWidth, circular: true });
      break;
    }
    case 'ellipse': {
      const [x, yBottom] = svgToPdfPoint(shape.cx - shape.rx, shape.cy + shape.ry, originX, originTopY, scale);
      flow.drawRect(x, yBottom, shape.rx * 2 * scale, shape.ry * 2 * scale, { fillRGB, strokeRGB, lineWidth, circular: true });
      break;
    }
    case 'line': {
      const [x1, y1] = svgToPdfPoint(shape.x1, shape.y1, originX, originTopY, scale);
      const [x2, y2] = svgToPdfPoint(shape.x2, shape.y2, originX, originTopY, scale);
      flow.drawLine(x1, y1, x2, y2, { strokeRGB, lineWidth });
      break;
    }
    case 'polygon':
    case 'polyline': {
      const points = shape.points.map(([x, y]) => svgToPdfPoint(x, y, originX, originTopY, scale));
      flow.drawPolygon(points, { fillRGB, strokeRGB, lineWidth, closed: shape.kind === 'polygon' });
      break;
    }
  }
}

function renderSvgShapes(flow, node) {
  flow.applySpacingBefore(0);
  let w = node.widthPx * PX_TO_PT;
  let h = node.heightPx * PX_TO_PT;

  const availableWidth = flow.contentWidth;
  const availableHeight = flow.remainingHeight() > 40 ? flow.remainingHeight() : flow.geometry.contentTop - flow.geometry.contentBottom;
  const scale = Math.min(1, availableWidth / w, availableHeight / h);
  w *= scale;
  h *= scale;

  flow.ensureSpace(h);
  const originX = flow.marginLeft + (availableWidth - w) / 2;
  const originTopY = flow.cursorY;
  // A single combined px→pt scale factor (the intrinsic px-to-pt
  // conversion times whatever additional shrink `scale` above applied to
  // fit the page) — aspect ratio is always preserved, so one factor covers
  // both X and Y for every shape's own coordinates.
  const unitScale = w / node.widthPx;

  for (const shape of node.shapes) drawSvgShape(flow, shape, originX, originTopY, unitScale);

  flow.advance(h);
}

// ─── Dispatch ───────────────────────────────────────────────────────────

export function renderBlock(flow, node, imageXObjectByAssetId) {
  switch (node.type) {
    case 'heading': return renderHeading(flow, node);
    case 'paragraph': return renderParagraph(flow, node);
    case 'quote': return renderQuote(flow, node, imageXObjectByAssetId);
    case 'code': return renderCode(flow, node);
    case 'list': return renderList(flow, node, 0, imageXObjectByAssetId);
    case 'thematicBreak': return renderThematicBreak(flow);
    case 'pageBreak': return flow.flushPage();
    case 'table': return renderTable(flow, node, imageXObjectByAssetId);
    case 'image': return renderImage(flow, node, imageXObjectByAssetId);
    case 'figure': {
      const uaFigure = uaDefaultMarginPt('figure');
      return renderImage(flow, node.image, imageXObjectByAssetId, node.caption,
        node.margin?.beforePt ?? uaFigure.beforePt, node.margin?.afterPt ?? uaFigure.afterPt);
    }
    case 'svgShapes': return renderSvgShapes(flow, node);
    case 'section': return renderBlockSequence(flow, node.children ?? [], imageXObjectByAssetId);
    default: return undefined;
  }
}

export { renderBlockSequence };
