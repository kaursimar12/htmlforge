/**
 * The pagination engine — the one genuinely new capability this exporter
 * needs that no other exporter in this project has. PPTX's own `bottomY`
 * (pptx/slide.js) is a soft per-block cap that just clips overflow onto a
 * fixed slide; DOCX relies entirely on Word's own viewer-side reflow (it
 * never generates more than one `w:sectPr`). PDF is fixed/rasterized at
 * generation time, so this module makes a real "does the next unit of
 * content fit on the current page? If not, flush it and start a new one"
 * decision — the first place in this codebase that does.
 */
import { ContentStreamBuilder, encodeWinAnsi } from './content-stream.js';
import { FontRegistry } from './fonts.js';
import { buildEmbeddedFontDescriptor } from './embedded-font.js';
import { resolveFontFamilyStack } from '../../fonts/resolver.js';
import { ROUNDED_CORNER_RADIUS_PT, DEFAULT_BODY_FONT_PT, DEFAULT_LINE_HEIGHT_MULTIPLIER } from './layout.js';

export class PdfPageFlow {
  #geometry;
  #bodyFontPt;
  #lineHeightMultiplier;
  #hasUnsupportedGlyphs = false;
  #fontRegistry = new FontRegistry();
  #imageEntryByAssetId = new Map();
  // Embedded (non-Standard-14) fonts — resolved lazily on first use during
  // rendering, the same "discover as blocks.js walks the document" pattern
  // `#fontRegistry.use()` already follows for standard fonts, rather than a
  // separate upfront IDM walk (images need that walk because asset bytes
  // must be fetched before rendering starts; a caller-supplied local font
  // file is cheap enough to parse on first encounter instead). Keyed by
  // "family|weight|italic" so requesting the same run style twice reuses one
  // resource name and one built descriptor. A `null` value marks a
  // family/weight/italic combo whose descriptor build already failed once
  // (see buildEmbeddedFontDescriptor's own warn-and-return-undefined
  // contract) — recorded so a second occurrence of the same failing run
  // style doesn't re-attempt (and re-warn about) the same doomed embed.
  #embeddedFontByKey = new Map();
  #embeddedFontRegistry;
  #warn;
  #pages = [];
  #current;
  // A temporary horizontal-slot override — set while rendering the content
  // of one grid/flex column, so every block-render function's existing
  // `flow.marginLeft`/`flow.contentWidth` reads (already used for plain
  // page-width rendering) transparently narrow to that column's own x/width
  // instead, with no separate "narrow" code path needed in blocks.js.
  #slot = null;
  // Real CSS collapses two adjoining block-level margins into the LARGER
  // of the two, not their sum (e.g. a heading's margin-bottom:16px flowing
  // into a paragraph's margin-top:16px leaves one 16pt gap, not 32pt) —
  // without this, every pair of adjacent blocks that both declare vertical
  // spacing render roughly twice the intended gap. `deferSpacingAfter`
  // registers what a block *wants* below it without moving the cursor yet;
  // the next block's own `applySpacingBefore` resolves the actual gap as
  // the max of what's pending and what it itself wants above it, then
  // clears the pending amount — so it's applied exactly once, not twice.
  #pendingSpacingPt = 0;
  cursorY;

  constructor(geometry, textDefaults = {}) {
    this.#geometry = geometry;
    // The document's own `body{font-size|line-height}` (see dom-to-idm.js's
    // metadata.fontSizePt/lineHeight) — falling back to this project's own
    // Normal-style-equivalent default only when the page never declared
    // one. Read once here and exposed to every block-render function via
    // `flow.bodyFontPt`/`flow.lineHeightMultiplier` (the same "carry it on
    // the one object every render call already receives" approach
    // `withHorizontalSlot`'s own contentWidth/marginLeft override uses),
    // rather than each of blocks.js's many render functions needing its own
    // extra parameter threaded through.
    this.#bodyFontPt = textDefaults.bodyFontPt ?? DEFAULT_BODY_FONT_PT;
    this.#lineHeightMultiplier = textDefaults.lineHeightMultiplier ?? DEFAULT_LINE_HEIGHT_MULTIPLIER;
    this.#embeddedFontRegistry = textDefaults.embeddedFontRegistry;
    this.#warn = textDefaults.warn;
    this.#startNewPage();
  }

  get geometry() { return this.#geometry; }
  get bodyFontPt() { return this.#bodyFontPt; }
  get lineHeightMultiplier() { return this.#lineHeightMultiplier; }
  get contentWidth() { return this.#slot ? this.#slot.width : this.#geometry.contentWidth; }
  get marginLeft() { return this.#slot ? this.#slot.left : this.#geometry.left; }
  /** 0-based index of the page currently being drawn to — table.js compares
   * this before/after drawing a row to detect a page break happened mid-table. */
  get pageIndex() { return this.#pages.length - 1; }

  /** Runs `fn` with `contentWidth`/`marginLeft` narrowed to one grid/flex
   * column's own slot — restored afterward even if `fn` throws. Slots
   * nest (a grid inside a grid cell narrows further from whatever slot, if
   * any, is already active), matching real CSS containment. */
  withHorizontalSlot(left, width, fn) {
    const previous = this.#slot;
    this.#slot = { left, width };
    try {
      return fn();
    } finally {
      this.#slot = previous;
    }
  }

  #startNewPage() {
    this.#current = {
      width: this.#geometry.pageWidth,
      height: this.#geometry.pageHeight,
      builder: new ContentStreamBuilder(),
      usedImageResourceNames: new Set(),
      links: [],
    };
    this.#pages.push(this.#current);
    this.cursorY = this.#geometry.contentTop;
  }

  /** Unconditional new page — used for an explicit PageBreakNode. */
  flushPage() {
    this.#startNewPage();
  }

  /** How much vertical room is left on the current page, in points. */
  remainingHeight() {
    return this.cursorY - this.#geometry.contentBottom;
  }

  /**
   * If `heightPt` doesn't fit in the remaining space, starts a new page
   * first. Called at the finest natural unit each block type already loops
   * over (one line of text, one table row, one list item) — see blocks.js —
   * so "a block taller than a page" is a non-event rather than a special
   * case: the loop just keeps drawing across however many pages it takes.
   */
  ensureSpace(heightPt) {
    if (this.cursorY - heightPt < this.#geometry.contentBottom) this.#startNewPage();
  }

  /** Moves the cursor down without drawing anything (inter-block spacing). */
  advance(heightPt) {
    this.cursorY -= heightPt;
  }

  /** Call before drawing a block that wants `desiredBeforePt` of space above
   * it — collapses with whatever a previous block already deferred (CSS
   * adjacent-margin collapsing: the effective gap is the larger of the two,
   * not the sum), applies that gap once, and clears the pending amount. */
  applySpacingBefore(desiredBeforePt) {
    const gap = Math.max(this.#pendingSpacingPt, desiredBeforePt);
    if (gap) this.advance(gap);
    this.#pendingSpacingPt = 0;
  }

  /** Call right after drawing a block to register how much trailing space
   * it wants — NOT applied immediately; the next block's own
   * `applySpacingBefore` decides the real gap. A block with nothing after
   * it (the last one in the document) never has its deferred space applied
   * at all, which is correct — trailing whitespace at the end of a page has
   * no visible effect anyway. */
  deferSpacingAfter(desiredAfterPt) {
    this.#pendingSpacingPt = Math.max(this.#pendingSpacingPt, desiredAfterPt);
  }

  /** Discards any deferred trailing spacing without applying it — used when
   * starting an independent vertical flow (a new grid/flex column, reset
   * to the row's own top Y) that must not inherit whatever the previous
   * column happened to leave pending; they're side by side, not stacked. */
  resetPendingSpacing() {
    this.#pendingSpacingPt = 0;
  }

  useFont(baseFont) {
    return this.#fontRegistry.use(baseFont);
  }

  /** Resolves a run's CSS `font-family` stack + weight/italic against the
   * caller-registered embedded fonts (see fonts/resolver.js) plus the
   * generic-keyword/web-safe/fallback waterfall — blocks.js's drawSegment
   * calls this before deciding whether to draw via `drawText` (Standard-14)
   * or `drawEmbeddedText` (a real embedded font). */
  resolveTextFont(fontFamilyCss, { weight = 400, italic = false } = {}) {
    return resolveFontFamilyStack(fontFamilyCss, { weight, italic }, this.#embeddedFontRegistry);
  }

  /**
   * Builds (on first use) or reuses the PDF resource description for one
   * embedded-font registry entry — deduped by family/weight/italic so a
   * repeated run style reuses one resource name and one built descriptor,
   * mirroring `useFont`'s own dedup for standard fonts. Returns `undefined`
   * if the descriptor build already failed once for this exact combo (see
   * `#embeddedFontByKey`'s own doc comment) — callers fall back to
   * `resolveTextFont`'s non-embedded result in that case, never throwing.
   */
  useEmbeddedFont(entry) {
    const key = `${entry.family}|${entry.weight}|${entry.italic}`;
    if (!this.#embeddedFontByKey.has(key)) {
      const resourceName = `EF${this.#embeddedFontByKey.size + 1}`;
      const descriptor = buildEmbeddedFontDescriptor(entry, resourceName, this.#warn);
      this.#embeddedFontByKey.set(key, descriptor ? { resourceName, descriptor } : null);
    }
    return this.#embeddedFontByKey.get(key) ?? undefined;
  }

  /**
   * Registers an already-built image XObject description (from
   * `images.js`'s `buildImageXObject`) for reuse across every page that
   * draws it, deduped by the source `assetId` — mirrors DOCX/PPTX's own
   * `imagesByAssetId`/`assetInfoByAssetId` maps.
   */
  registerImage(assetId, xObject) {
    if (!this.#imageEntryByAssetId.has(assetId)) {
      this.#imageEntryByAssetId.set(assetId, xObject);
    }
    return this.#imageEntryByAssetId.get(assetId);
  }

  /** `y` is the text baseline, in PDF page-space (bottom-up), not the top of the glyph box. */
  drawText(x, y, text, { baseFont, sizePt, colorRGB = [0, 0, 0] }) {
    if (!text) return;
    const { text: encoded, hasUnsupported } = encodeWinAnsi(text);
    if (hasUnsupported) this.#hasUnsupportedGlyphs = true;
    const resourceName = this.useFont(baseFont);
    const b = this.#current.builder;
    b.save();
    b.setFillColorRGB(...colorRGB);
    b.beginText();
    b.setFont(resourceName, sizePt);
    b.setTextPosition(x, y);
    b.showText(encoded);
    b.endText();
    b.restore();
  }

  /** The embedded-font sibling of `drawText` — `resourceName`/`descriptor`
   * come from `useEmbeddedFont`. Each character maps to that font's own
   * glyph ID via its `cmap` (see embedded-font.js's `glyphIdForCodepoint`);
   * a codepoint the font doesn't cover at all falls back to glyph 0
   * (`.notdef`, conventionally a blank box or missing-glyph marker — the
   * same "can't render this exact character, don't crash" spirit as
   * `encodeWinAnsi`'s own literal `?` substitution) and is counted toward
   * the same document-wide `hasUnsupportedGlyphs` warning. Glyph IDs are
   * shown as a `<...>` hex string, 2 bytes (4 hex digits) each — see
   * `content-stream.js`'s `showHexText` for why this needs its own encoding
   * path instead of `drawText`'s single-byte WinAnsi one. */
  drawEmbeddedText(x, y, text, { resourceName, descriptor, sizePt, colorRGB = [0, 0, 0] }) {
    if (!text) return;
    let hex = '';
    for (const ch of text) {
      const codePoint = ch.codePointAt(0);
      const glyphId = descriptor.glyphIdForCodepoint(codePoint);
      if (glyphId === undefined) this.#hasUnsupportedGlyphs = true;
      hex += (glyphId ?? 0).toString(16).padStart(4, '0');
    }
    const b = this.#current.builder;
    b.save();
    b.setFillColorRGB(...colorRGB);
    b.beginText();
    b.setFont(resourceName, sizePt);
    b.setTextPosition(x, y);
    b.showHexText(hex);
    b.endText();
    b.restore();
  }

  /** Whether any drawn text contained a character with no WinAnsiEncoding
   * representation (CJK/Arabic/Cyrillic/emoji) — `index.js` surfaces this
   * as one warning for the whole document, not per character. */
  get hasUnsupportedGlyphs() { return this.#hasUnsupportedGlyphs; }

  /** `rounded`/`circular` pick a rounded-rect or ellipse path instead of a
   * plain rectangle — the PDF-native equivalent of DrawingML's `roundRect`/
   * `ellipse` preset geometry (see layout.js's ROUNDED_CORNER_RADIUS_PT). */
  drawRect(x, y, w, h, { fillRGB, strokeRGB, lineWidth = 1, rounded = false, circular = false } = {}) {
    const b = this.#current.builder;
    if (!fillRGB && !strokeRGB) return;
    b.save();
    if (circular) b.ellipse(x, y, w, h);
    else if (rounded) b.roundedRect(x, y, w, h, Math.min(ROUNDED_CORNER_RADIUS_PT, w / 2, h / 2));
    else b.rect(x, y, w, h);
    if (fillRGB && strokeRGB) {
      b.setFillColorRGB(...fillRGB);
      b.setStrokeColorRGB(...strokeRGB);
      b.setLineWidth(lineWidth);
      b.fillAndStroke();
    } else if (fillRGB) {
      b.setFillColorRGB(...fillRGB);
      b.fill();
    } else {
      b.setStrokeColorRGB(...strokeRGB);
      b.setLineWidth(lineWidth);
      b.stroke();
    }
    b.restore();
  }

  drawLine(x1, y1, x2, y2, { strokeRGB = [0, 0, 0], lineWidth = 1 } = {}) {
    const b = this.#current.builder;
    b.save();
    b.setStrokeColorRGB(...strokeRGB);
    b.setLineWidth(lineWidth);
    b.moveTo(x1, y1);
    b.lineTo(x2, y2);
    b.stroke();
    b.restore();
  }

  /** An arbitrary multi-point shape (SVG's `<polygon>`/`<polyline>`).
   * `closed` explicitly draws the closing edge (`h`) before painting — a
   * polygon always closes (SVG auto-closes it), a polyline doesn't (its
   * stroke has no segment back to the start point). Either way, `f`/`B`
   * still fills the path as if closed regardless of `closed` — that's
   * PDF's own native fill semantics for open subpaths (PDF spec §8.5.3),
   * which happens to already match SVG's own polyline fill behavior
   * exactly, with no special-casing needed here. */
  drawPolygon(points, { fillRGB, strokeRGB, lineWidth = 1, closed = true } = {}) {
    if (points.length < 2 || (!fillRGB && !strokeRGB)) return;
    const b = this.#current.builder;
    b.save();
    b.moveTo(points[0][0], points[0][1]);
    for (let i = 1; i < points.length; i++) b.lineTo(points[i][0], points[i][1]);
    if (closed) b.closePath();
    if (fillRGB && strokeRGB) {
      b.setFillColorRGB(...fillRGB);
      b.setStrokeColorRGB(...strokeRGB);
      b.setLineWidth(lineWidth);
      b.fillAndStroke();
    } else if (fillRGB) {
      b.setFillColorRGB(...fillRGB);
      b.fill();
    } else {
      b.setStrokeColorRGB(...strokeRGB);
      b.setLineWidth(lineWidth);
      b.stroke();
    }
    b.restore();
  }

  /** An image XObject paints into the unit square (0,0)-(1,1) — one combined
   * scale+translate `cm` matrix maps that square to the real drawn box
   * before `Do` paints it. */
  drawImage(xObject, x, y, w, h) {
    this.#current.usedImageResourceNames.add(xObject.resourceName);
    const b = this.#current.builder;
    b.save();
    b.concatMatrix(w, 0, 0, h, x, y);
    b.drawImage(xObject.resourceName);
    b.restore();
  }

  addLinkAnnotation(rect, uri) {
    this.#current.links.push({ rect, uri });
  }

  finish() {
    return {
      pages: this.#pages.map((p) => ({
        width: p.width,
        height: p.height,
        ops: p.builder.toString(),
        usedImageResourceNames: p.usedImageResourceNames,
        links: p.links,
      })),
      fontEntries: this.#fontRegistry.entries(),
      imageEntries: [...this.#imageEntryByAssetId.values()],
      // Only successfully-built descriptors — a combo whose embed failed
      // (recorded as `null` in #embeddedFontByKey, see useEmbeddedFont) has
      // nothing for pdf-writer.js to allocate objects for; its runs already
      // fell back to a Standard-14 draw instead.
      embeddedFontEntries: [...this.#embeddedFontByKey.values()].filter(Boolean).map((v) => v.descriptor),
    };
  }
}
