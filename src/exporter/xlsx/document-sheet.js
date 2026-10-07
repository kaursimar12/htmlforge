/**
 * Renders an entire document — headings, paragraphs, lists, quotes, code
 * blocks, and tables — into ONE worksheet, stacked in document order, instead
 * of the previous "one worksheet per <table>" model. Approximates HTML's
 * visual page: a heading/paragraph becomes a full-width row (background
 * fill/bold/alignment carried over from its own IDM fields, via the same
 * run-flattening + value-classification path a table cell's content already
 * goes through), a table renders its own grid at whatever row it's reached,
 * and consecutive blocks that don't share a background color get a blank
 * spacer row between them.
 *
 * Doesn't attempt real CSS layout (floats/flex/absolute positioning) — see
 * collect-document-flow.js's doc comment. Every non-table block always spans
 * the sheet's full column width; only a table's own column offset (see
 * computeTableStartCol below) responds to `margin-left:auto`, and its column
 * widths (see the widthPercent scaling in documentSheetXml) to `width:%`.
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';
import { layoutGrid, renderTableGrid, computeColumnWidths } from './worksheet.js';
import { flattenBlockRuns, trimRuns } from './rich-text.js';
import { cellRef, classifyCellValue } from './cell-value.js';
import { resolveImagePxSize } from '../shared/image-size.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
// Only needed for the sheet's own <drawing r:id="..."/> reference (X8,
// images) — see the caller (xlsx/index.js), which splices that element in
// only when the document actually has an embedded image.
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Same category of "no real layout engine, so a plausible default" as
// worksheet.js's MIN/MAX_COL_WIDTH — used to size the sheet's column count
// when there's no table at all to measure (a bare heading/paragraph has no
// natural column count of its own), and as the starting width for any column
// no table ever touches.
const DEFAULT_COLUMN_COUNT = 4;
const DEFAULT_COL_WIDTH = 12;

// A text row's default font size when no run carries its own — matches
// Calibri 11's default already assumed elsewhere in this exporter (styles.js).
const DEFAULT_FONT_SIZE_PT = 11;
// Real spreadsheet apps give a single line of text noticeably more vertical
// room than a plain 1.2x leading: Excel's own default row height for 11pt
// Calibri is ~15pt (a ~1.36x factor), Google Sheets' default row is 21px
// (~1.575x for its own default font). 1.2x under-estimates either one — fine
// for a single-line row (Excel/Sheets still show one full line even if the
// row is a touch short), but for a MULTI-line block (several `<br>`-joined
// lines in one wrapped cell, e.g. an address block) that per-line shortfall
// compounds across every line, visibly clipping/overflowing the last one or
// two lines into the row below (observed: a 3-line block already noticeably
// short, a 6-line block cut off by roughly half a line's worth more). 1.5x
// sits above both real apps' own single-line defaults, so a wrapped
// multi-line block reliably gets enough room on either.
const LINE_HEIGHT_MULTIPLIER = 1.5;

// Same 96px/in reference every other px->pt conversion in this exporter uses
// (see worksheet.js's identical POINTS_PER_PX) — used to convert a cardGroup
// band's explicit CSS `height` (dom-to-idm.js's heightPx) into points before
// comparing it against the band's own accumulated row heights.
const PX_TO_PT = 0.75;

function tableColumnCount(table) {
  return layoutGrid([...(table.head ?? []), ...table.rows]).columnCount;
}

// A table's own column COUNT is always contiguous and exactly as wide as its
// own grid — `widthPercent` never pads it with extra blank columns (that left
// a visible gap between the table and the sheet's right edge); it instead
// scales the table's own column WIDTHS (see documentSheetXml below). Absent
// `marginAuto`, a table just hugs the left edge at its natural width.
function computeTableStartCol(table, ownColumnCount, sheetColumnCount) {
  if (table.marginAuto === 'left') return Math.max(0, sheetColumnCount - ownColumnCount); // hugs the right edge
  if (table.marginAuto === 'both') return Math.max(0, Math.floor((sheetColumnCount - ownColumnCount) / 2)); // centered
  return 0; // hugs the left edge (marginAuto === 'right', or no signal at all)
}

// Splits `targetCols` shared grid columns across `weights.length` of a
// narrower table's own columns, proportional to each column's own relative
// weight (its natural content-driven width) — largest-remainder rounding so
// the shares always sum to EXACTLY targetCols, not just approximately, and
// every column gets at least 1 so none disappears entirely.
function distributeColumns(weights, targetCols) {
  const total = weights.reduce((a, w) => a + w, 0);
  const raw = weights.map((w) => (total > 0 ? (w / total) * targetCols : targetCols / weights.length));
  const spans = raw.map((v) => Math.max(1, Math.floor(v)));
  const order = raw.map((v, i) => ({ i, frac: v - Math.floor(v) })).sort((a, b) => b.frac - a.frac);
  let used = spans.reduce((a, b) => a + b, 0);
  for (let k = 0; used < targetCols; k++) {
    spans[order[k % order.length].i] += 1;
    used += 1;
  }
  while (used > targetCols) {
    const maxIdx = spans.reduce((best, v, i) => (v > spans[best] ? i : best), 0);
    if (spans[maxIdx] <= 1) break; // every column already at the 1-column floor — can't shrink further
    spans[maxIdx] -= 1;
    used -= 1;
  }
  return spans;
}

// A table meant to span the FULL sheet width (`width:100%`, or no width
// declared at all — the common case for a plain full-bleed data table) but
// with FEWER of its own `<td>` columns than a sibling table elsewhere in the
// same document (e.g. this project's own 2-column "From"/"Bill To" billing
// table sitting above a 5-column items table) used to just scale its own 2
// columns' raw WIDTHS up to "100% of the sheet's width so far" (see
// documentSheetXml's own widthPercent-scaling comment) — visually correct in
// isolation, but it meant every OTHER full-width table with a different
// native column count ended up on a completely different column grid, so
// their cell boundaries never lined up with each other at all (an "Invoice
// Items" table's own 5 real columns vs. a "From"/"Bill To" table's own 2,
// each independently stretched to "the full width" by width alone, with no
// shared column boundaries between them). Remapping this table's own cells
// via colspan onto the SAME shared `targetColumnCount` every other full-width
// table in the document also uses — proportional to each of its own
// columns' natural content width, via distributeColumns above — is what
// actually lines every full-width table's cell edges up with each other,
// the same way they'd align in a browser (where every one of these tables
// independently stretches to the same page width regardless of its own
// column count). Only used for a table that's genuinely trying to span the
// FULL width; a deliberately narrower table (e.g. a `width:50%` totals
// summary) keeps using the existing width-scaling approach instead — see its
// own call site in documentSheetXml.
function remapTableToColumnCount(table, targetColumnCount) {
  const headRows = table.head ?? [];
  const allRows = [...headRows, ...table.rows];
  if (allRows.length === 0) return table;
  const { grid, columnCount: ownColumnCount } = layoutGrid(allRows);
  if (ownColumnCount >= targetColumnCount) return table;

  const weights = computeColumnWidths(grid, ownColumnCount);
  const spans = distributeColumns(weights, targetColumnCount);

  const remappedRows = grid.map((gridRow, rowIndex) => {
    const cells = [...gridRow.entries()].sort(([a], [b]) => a - b).map(([col, cell]) => {
      const span = spans.slice(col, col + (cell.colspan ?? 1)).reduce((a, b) => a + b, 0);
      return { ...cell, colspan: span > 1 ? span : undefined };
    });
    return { ...allRows[rowIndex], cells };
  });

  return {
    ...table,
    head: headRows.length ? remappedRows.slice(0, headRows.length) : table.head,
    rows: remappedRows.slice(headRows.length),
  };
}

// One heading/paragraph/list/quote/code block -> one full-width row.
// `extraHtPt` (see documentSheetXml's cardGroup band tracking) pads this
// row's own font-driven height to make the *combined* height of every row in
// a shared-background band reach an explicit CSS `height` on the `<div>`
// that band collapsed from — added only to the band's last row, the same
// "slack absorbed at the end" convention the PDF exporter's renderCardGroup
// uses for the same underlying gap (a spreadsheet row has no equivalent of
// vertical-align:top text sitting inside a taller box; only the row height
// itself can grow).
function renderBlockRow(node, rowIndex, columnCount, rctx, extraHtPt = 0) {
  const runs = trimRuns(flattenBlockRuns(node));
  const text = runs.map((r) => r.text).join('');
  const fontStyle = runs.length === 1 ? runs[0] : {};
  const value = classifyCellValue(text);
  const lines = text.split('\n').length;

  const styleIndex = rctx.styles.registerCellStyle({
    ...fontStyle,
    backgroundColor: node.backgroundColor,
    align: node.align,
    // Excel's own default (no vertical alignment set at all) is BOTTOM, not
    // middle — invisible for a normal-height row, but these rows are often
    // made much taller than one line by padding/font size (see above), and a
    // bottom-stuck single line in a tall colored band looks visibly wrong
    // next to the browser's naturally centered box content. Only defaulted
    // when there's actual content — a genuinely empty, unstyled block must
    // stay on the shared default style (index 0) to keep omitting its own
    // <c> cells entirely (see the value.type === 'empty' branch below).
    verticalAlign: value.type === 'empty' ? undefined : 'center',
    numFmtCode: value.type === 'number' ? value.numFmtCode : undefined,
    wrapText: value.type !== 'number' && lines > 1,
  });

  // Sized to the content's own font (the largest run's, e.g. a heading in a
  // mixed-run block) rather than a flat guess — otherwise a large heading
  // font gets clipped into a normal-height row while small-font body text
  // gets an oversized one (both looked visibly wrong against the source HTML)
  // — plus the block's own top+bottom padding (e.g. a `padding:20px` banner
  // paragraph), the same "invisible to Excel unless added explicitly" gap
  // worksheet.js's estimateRowHeightPt fixes for table cells.
  const paddingTopPt = node.padding?.top ?? 0;
  const paddingBottomPt = node.padding?.bottom ?? 0;
  const rowAttrs = { r: rowIndex + 1 };
  const hasOwnHeight = text !== '' || paddingTopPt > 0 || paddingBottomPt > 0;
  const maxFontSizePt = Math.max(rctx.defaultFontSizePt ?? DEFAULT_FONT_SIZE_PT, ...runs.map((r) => r.fontSizePt ?? 0));
  const baseHt = hasOwnHeight ? lines * maxFontSizePt * LINE_HEIGHT_MULTIPLIER + paddingTopPt + paddingBottomPt : 0;
  if (hasOwnHeight || extraHtPt > 0) {
    rowAttrs.ht = baseHt + extraHtPt;
    rowAttrs.customHeight = 1;
  }

  const b = new XmlBuilder();
  b.el('row', rowAttrs);

  if (value.type === 'empty') {
    if (styleIndex !== 0) {
      for (let col = 0; col < columnCount; col++) b.self('c', { r: cellRef(rowIndex, col), s: styleIndex });
    }
  } else {
    const ref = cellRef(rowIndex, 0);
    if (value.type === 'number') {
      b.el('c', styleIndex !== 0 ? { r: ref, s: styleIndex } : { r: ref });
      b.el('v').text(value.value).up();
      b.up();
    } else {
      const stringIndex = runs.length > 1 ? rctx.sharedStrings.addRich(runs) : rctx.sharedStrings.add(text);
      b.el('c', styleIndex !== 0 ? { r: ref, t: 's', s: styleIndex } : { r: ref, t: 's' });
      b.el('v').text(String(stringIndex)).up();
      b.up();
    }
    for (let col = 1; col < columnCount; col++) {
      if (styleIndex !== 0) b.self('c', { r: cellRef(rowIndex, col), s: styleIndex });
    }
  }

  b.up();
  return { xml: b.end(), merge: columnCount > 1 ? { r1: rowIndex, c1: 0, r2: rowIndex, c2: columnCount - 1 } : null, heightPt: baseHt };
}

// X8 — an image/figure block reserves one row sized to the image's own pixel
// height, with the image itself anchored via a real xl/drawings part (see
// xlsx/drawing.js's drawingXml) rather than any cell content. A figure's
// caption (if any) renders as a normal text row directly below — via the
// existing renderBlockRow, on a synthesized italic paragraph — the same
// image-then-caption stacking DOCX's own renderFigure uses.
function renderImageRow(node, rowIndex, columnCount, rctx) {
  const imageNode = node.type === 'figure' ? node.image : node;
  const asset = rctx.images?.get(imageNode.assetId);
  const rowsXml = [];
  let row = rowIndex;

  // Unresolved (fetch failed, no assetId, or embedImages: false) images
  // reserve no row at all — resolveAssets already emitted a ctx.warn for the
  // failure, matching DOCX/PPTX/PDF's identical silent-skip convention.
  if (asset) {
    const { widthPx, heightPx } = resolveImagePxSize(imageNode, asset);
    const b = new XmlBuilder();
    b.self('row', { r: row + 1, ht: heightPx * PX_TO_PT, customHeight: 1 });
    rowsXml.push(b.end());
    rctx.anchors.push({ row, col: 0, widthPx, heightPx, relId: asset.relId, filename: asset.filename });
    row += 1;
  }

  if (node.type === 'figure' && node.caption) {
    const caption = { type: 'paragraph', children: [{ type: 'emphasis', children: [{ type: 'text', value: node.caption }] }] };
    const { xml } = renderBlockRow(caption, row, columnCount, rctx);
    rowsXml.push(xml);
    row += 1;
  }

  return { rowsXml, rowCount: row - rowIndex };
}

// A blank row between two blocks — skipped only when both sit inside the
// same visually-continuous background band (approximates HTML's padding/
// margin gaps without converting them to precise row heights).
function needsSpacer(prevNode, node) {
  if (!prevNode) return false;
  if (prevNode.type === 'table' || node.type === 'table') return true;
  return !(prevNode.backgroundColor !== undefined && prevNode.backgroundColor === node.backgroundColor);
}

// Excel's own column-width unit is "characters at the workbook's default
// font's max digit width" — for this exporter's assumed Calibri 11 default
// (styles.js), that digit width is ~7px, so a pixel measurement divides by 7
// to land in the same unit colWidths already uses everywhere else (the
// `len + 2` character-count estimate in worksheet.js's computeColumnWidths).
// Approximate, not exact (Excel's real formula also has a small additive
// padding term) — same "good enough, not pixel-perfect" precision every
// other px conversion in this exporter already accepts.
const PX_PER_COL_WIDTH_UNIT = 7;

/**
 * `blocks` is collectDocumentFlow's output; `rctx` is `{ sharedStrings, styles
 * [, images] }` — `images` (X8, a `Map<assetId, {relId, filename}>` built by
 * xlsx/index.js from resolveAssets' own output, same shape DOCX/PPTX already
 * use) is only needed when the document has images/figures. Returns `{ xml,
 * anchors }` — `anchors` (one entry per successfully embedded image, empty
 * when there are none) is what xlsx/index.js uses to build xl/drawings/
 * drawing1.xml; the sheet's own `<drawing r:id="...">` reference is spliced
 * into `xml` by the caller once it knows whether that part exists at all.
 */
export function documentSheetXml(blocks, rctx, contentWidthPx) {
  rctx.anchors = [];
  const tables = blocks.filter((b) => b.type === 'table');
  // Distinct from `columnCount` just below: that one is floored to
  // DEFAULT_COLUMN_COUNT purely to give the SHEET itself a sane minimum
  // width when there's no table at all (or only narrow ones) to measure —
  // it is NOT a real table's own column count, so remapping a genuinely
  // 2-column document (every table narrower than the default floor) up to
  // that arbitrary 4 would widen it for no reason. The full-width colspan
  // remap below (see remapTableToColumnCount) only ever targets the widest
  // REAL table actually present in the document.
  const widestTableColumnCount = Math.max(0, ...tables.map(tableColumnCount));
  const columnCount = Math.max(DEFAULT_COLUMN_COUNT, widestTableColumnCount);

  const rowsXml = [];
  const merges = [];
  const colWidths = new Array(columnCount).fill(DEFAULT_COL_WIDTH);
  let currentRow = 0;
  let prevNode = null;
  // Tracks the running total of already-rendered row heights across a run of
  // consecutive blocks sharing one `cardGroup` (dom-to-idm.js's shared token
  // for every member of a `<div>` that collapsed into several rows here) —
  // reset whenever a new group starts, or for a standalone (cardGroup-less)
  // block, every single row. Lets the band's LAST row absorb whatever slack
  // remains between the group's own accumulated row heights and its
  // explicit CSS `height`, the same "pad the end, not each row" convention
  // the PDF exporter's renderCardGroup uses for the identical gap.
  let cardGroupAccumHt = 0;
  let cardGroupRef = undefined;

  for (let i = 0; i < blocks.length; i++) {
    const node = blocks[i];
    if (needsSpacer(prevNode, node)) currentRow += 1;

    if (node.type === 'table') {
      cardGroupAccumHt = 0;
      cardGroupRef = undefined;
      const nativeColumnCount = tableColumnCount(node);
      // A table meant to span the FULL sheet width (no declared width, or
      // effectively 100%) but built from fewer native <td> columns than some
      // OTHER full-width table elsewhere in the same document (this
      // project's own 2-column "From"/"Bill To" billing table above a
      // 5-column items table, both `width:100%`) gets its own columns
      // remapped via colspan onto the SAME shared columnCount every
      // full-width table uses — see remapTableToColumnCount's own note on
      // why that's what actually lines their cell edges up with each other.
      // A genuinely narrower table (e.g. a `width:50%` totals summary) is
      // untouched here and instead widened by the existing per-column width
      // scaling just below — remapping via colspan only makes sense for a
      // table trying to claim the FULL width, not a deliberately partial one.
      const isFullWidth = node.widthPercent == null || node.widthPercent >= 99;
      const needsRemap = isFullWidth && nativeColumnCount < widestTableColumnCount;
      const renderableTable = needsRemap ? remapTableToColumnCount(node, widestTableColumnCount) : node;
      const ownColumnCount = needsRemap ? widestTableColumnCount : nativeColumnCount;
      const startCol = computeTableStartCol(renderableTable, ownColumnCount, columnCount);

      const grid = renderTableGrid(renderableTable, rctx, { startRow: currentRow, startCol });
      for (const row of grid.rows) rowsXml.push(row.xml);
      merges.push(...grid.merges);

      // `width:%` scales this table's own column widths up (never down) so
      // together they visually claim roughly that share of the sheet's
      // established width so far — e.g. a `width:50%` totals table renders
      // with noticeably wider columns than its own short "Total"/"$1,260"
      // text would otherwise justify, matching how it reads in the browser.
      // Only meaningful when the table occupies fewer columns than the sheet
      // has — a table already spanning every column (e.g. the common
      // `width:100%` on a full-width data table) has nothing to scale
      // against; skipping it there avoids needlessly inflating that table's
      // own columns off of the sheet's (somewhat arbitrary) starting default.
      let widths = grid.colWidths;
      if (node.widthPercent != null && ownColumnCount < columnCount) {
        const referenceTotal = colWidths.reduce((sum, w) => sum + w, 0);
        const naturalTotal = widths.reduce((sum, w) => sum + w, 0);
        const targetTotal = (node.widthPercent / 100) * referenceTotal;
        if (naturalTotal > 0 && targetTotal > naturalTotal) {
          const scale = targetTotal / naturalTotal;
          widths = widths.map((w) => w * scale);
        }
      }
      widths.forEach((w, idx) => {
        const col = startCol + idx;
        if (col < columnCount) colWidths[col] = Math.max(colWidths[col], w);
      });
      currentRow += grid.rows.length;
    } else if (node.type === 'image' || node.type === 'figure') {
      cardGroupAccumHt = 0;
      cardGroupRef = undefined;
      const { rowsXml: imageRows, rowCount } = renderImageRow(node, currentRow, columnCount, rctx);
      rowsXml.push(...imageRows);
      currentRow += rowCount;
    } else {
      if (node.cardGroup === undefined || node.cardGroup !== cardGroupRef) {
        cardGroupAccumHt = 0;
        cardGroupRef = node.cardGroup;
      }

      const rendered = renderBlockRow(node, currentRow, columnCount, rctx);
      cardGroupAccumHt += rendered.heightPt;

      // Only the band's LAST row (a standalone block is always its own
      // one-row "band") gets padded — mid-band rows keep their own plain
      // font-driven height, matching every other exporter's "slack goes at
      // the end" convention.
      const isLastOfBand = node.cardGroup === undefined || blocks[i + 1]?.cardGroup !== node.cardGroup;
      let { xml, merge } = rendered;
      if (isLastOfBand && node.heightPx != null) {
        const explicitHeightPt = node.heightPx * PX_TO_PT;
        const deficitPt = explicitHeightPt - cardGroupAccumHt;
        if (deficitPt > 0) ({ xml, merge } = renderBlockRow(node, currentRow, columnCount, rctx, deficitPt));
      }

      rowsXml.push(xml);
      if (merge) merges.push(merge);
      currentRow += 1;
    }

    prevNode = node;
  }

  // Every column width computed above is purely content-driven (auto-fit to
  // the longest cell text) — with no connection at all to the source page's
  // actual design width (dom-to-idm.js's own metadata.contentWidthPx, the
  // same signal the PDF exporter uses to size its page). A `width:100%`
  // table — the common case, spanning every sheet column — never triggers
  // the widthPercent scaling above (that only fires when a table occupies
  // FEWER columns than the sheet has), so without this, the sheet's total
  // width always renders far narrower than a source page designed at, say,
  // 1200px, no matter how wide that page's own table was meant to be. Scaled
  // uniformly (preserving each column's already-established relative
  // proportion, the same way a browser stretches an auto-layout table to
  // fill a wider container) and, like the widthPercent scaling above, only
  // ever scaled UP — a page whose natural content already exceeds its own
  // declared width keeps its natural (wider) sizing rather than being
  // squeezed down to match.
  if (contentWidthPx) {
    const targetTotal = contentWidthPx / PX_PER_COL_WIDTH_UNIT;
    const naturalTotal = colWidths.reduce((sum, w) => sum + w, 0);
    if (naturalTotal > 0 && targetTotal > naturalTotal) {
      const scale = targetTotal / naturalTotal;
      for (let i = 0; i < colWidths.length; i++) colWidths[i] *= scale;
    }
  }

  const b = new XmlBuilder();
  b.el('worksheet', { xmlns: NS_MAIN, 'xmlns:r': NS_R });

  b.el('cols');
  colWidths.forEach((width, i) => b.self('col', { min: i + 1, max: i + 1, width, customWidth: 1 }));
  b.up();

  b.el('sheetData');
  for (const xml of rowsXml) b.raw(xml);
  b.up();

  if (merges.length > 0) {
    b.el('mergeCells', { count: merges.length });
    for (const m of merges) b.self('mergeCell', { ref: `${cellRef(m.r1, m.c1)}:${cellRef(m.r2, m.c2)}` });
    b.up();
  }

  return { xml: xmlDoc(b.end()), anchors: rctx.anchors };
}
