/**
 * xl/worksheets/sheetN.xml — one table's grid of rows/cells, plus its
 * <cols> (X6, auto-width) and <mergeCells> (X7, colspan/rowspan) — plus
 * later additions: rich-text runs, vertical align, direction/writing-mode/
 * rotation (see rich-text.js/cellTextRotation), row height.
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';
import { flattenCellText, isBlockNode } from '../../idm/flatten-text.js';
import { cellRef, classifyCellValue } from './cell-value.js';
import { flattenCellRuns, trimRuns } from './rich-text.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

// A <td> holding several distinctly-background-colored block children (the
// common "colored label bar followed by plain address lines" pattern — e.g.
// `<td><h3 style="background:#1F4E79;color:white">From</h3><p>...</p>
// <p>...</p></td>`) can't have that per-block background rendered within ONE
// Excel cell — a cell's own fill is all-or-nothing for its whole area,
// unlike a browser's own per-element background. Splitting the row into a
// couple of dedicated physical rows (the heading gets its OWN row/cell with
// its OWN real fill, the plain lines get their own row below with no fill)
// is the only way to actually show that color banding in XLSX at all, short
// of either dropping the color entirely (plain, correct-but-flat text — see
// rich-text.js's own dropBackgroundPairedColor) or promoting it to the WHOLE
// cell's fill (visually wrong — the address lines would turn blue too).
// Grouped by backgroundColor value (`undefined` counts as its own group) so
// any number of alternating bands works, not just "heading then plain text"
// specifically. Returns null when there's nothing worth splitting — a
// single shared background (including none at all) across every child.
function splitIntoBackgroundBands(cell) {
  const children = cell.children ?? [];
  if (children.length < 2 || !children.every(isBlockNode)) return null;
  const bands = [];
  for (const child of children) {
    const backgroundColor = child.backgroundColor ?? cell.backgroundColor;
    const last = bands[bands.length - 1];
    if (last && last.backgroundColor === backgroundColor) last.children.push(child);
    else bands.push({ backgroundColor, children: [child] });
  }
  return bands.length > 1 ? bands : null;
}

// Expands one TableRowNode into several physical rows when any of its cells
// has multiple background bands (see splitIntoBackgroundBands above). A
// sibling cell that DOESN'T split (the common case: a plain data cell, or
// one with only a single background throughout) spans every physical row
// this logical row expanded into via a real rowSpan merge, so it still reads
// as one continuous box rather than being cut into empty-looking fragments.
// A cell with FEWER bands than this row's max only pads its OWN *last* band
// with that rowSpan (the same "slack absorbed at the end" convention already
// used throughout this project — see document-sheet.js's own cardGroup
// handling). border-top/-bottom are only kept on the first/last band
// respectively (left/right stay on every band, since column edges don't
// change when a row splits vertically) — otherwise every internal band
// boundary would draw a redundant double border line that the original
// single-box design never had. Padding's top/bottom split the same way, so
// the visual gap doesn't double up at the fold between two bands.
function expandRowForBands(row) {
  const perCellBands = row.cells.map((cell) => ({ cell, bands: splitIntoBackgroundBands(cell) }));
  const maxBands = Math.max(1, ...perCellBands.map(({ bands }) => bands?.length ?? 1));
  if (maxBands === 1) return [row];

  const physicalRows = [];
  for (let bandIndex = 0; bandIndex < maxBands; bandIndex++) {
    const cells = [];
    for (const { cell, bands } of perCellBands) {
      if (!bands) {
        if (bandIndex === 0) cells.push({ ...cell, rowspan: maxBands });
        continue;
      }
      if (bandIndex >= bands.length) continue;
      const band = bands[bandIndex];
      const isFirstBand = bandIndex === 0;
      const isLastBand = bandIndex === bands.length - 1;
      const span = isLastBand ? maxBands - bandIndex : 1;
      cells.push({
        ...cell,
        backgroundColor: band.backgroundColor,
        children: band.children,
        borderTop: isFirstBand ? cell.borderTop : undefined,
        borderBottom: isLastBand ? cell.borderBottom : undefined,
        padding: cell.padding && {
          ...cell.padding,
          top: isFirstBand ? cell.padding.top : 0,
          bottom: isLastBand ? cell.padding.bottom : 0,
        },
        rowspan: span > 1 ? span : undefined,
      });
    }
    // Only the first physical row can plausibly inherit the source <tr>'s
    // own explicit CSS height — the rest are new rows this expansion itself
    // introduced, with no explicit height of their own to carry.
    physicalRows.push({ ...row, cells, heightPx: bandIndex === 0 ? row.heightPx : undefined });
  }
  return physicalRows;
}

function expandRowsForBands(rows) {
  return rows.flatMap(expandRowForBands);
}

// Expands the IDM's collapsed rowspan/colspan rows into a real grid: each
// row is a Map<colIndex, cell> holding only the *origin* cell of each
// span — unlike DOCX's WordprocessingML vMerge model, SpreadsheetML doesn't
// need continuation placeholder cells for the rows/columns a span covers;
// Excel treats a missing cell reference as blank and <mergeCell> alone
// tells it to visually span that range. `pending` tracks which columns a
// still-open rowspan from an earlier row currently occupies, so later
// cells in the same row correctly land one column further right.
export function layoutGrid(rows) {
  const pending = new Map(); // col -> { rowsLeft, colspan }
  const grid = [];
  const merges = [];
  const heightsPx = [];
  let maxCols = 0;

  rows.forEach((row, rowIndex) => {
    heightsPx.push(row.heightPx);
    const gridRow = new Map();
    let col = 0;
    let cellIndex = 0;

    while (cellIndex < row.cells.length || [...pending.keys()].some((c) => c >= col)) {
      const p = pending.get(col);
      if (p && p.rowsLeft > 0) {
        p.rowsLeft--;
        if (p.rowsLeft === 0) pending.delete(col);
        col += p.colspan;
        continue;
      }
      if (cellIndex >= row.cells.length) {
        const next = [...pending.keys()].filter((c) => c >= col).sort((a, b) => a - b)[0];
        if (next === undefined) break;
        col = next;
        continue;
      }
      const cell = row.cells[cellIndex++];
      const colspan = cell.colspan ?? 1;
      const rowspan = cell.rowspan ?? 1;
      gridRow.set(col, cell);
      if (colspan > 1 || rowspan > 1) {
        merges.push({ r1: rowIndex, c1: col, r2: rowIndex + rowspan - 1, c2: col + colspan - 1, cell });
      }
      if (rowspan > 1) pending.set(col, { rowsLeft: rowspan - 1, colspan });
      col += colspan;
    }

    maxCols = Math.max(maxCols, col);
    grid.push(gridRow);
  });

  return { grid, merges, heightsPx, columnCount: maxCols || 1 };
}

// Excel row height is expressed in points, not pixels — 1px = 0.75pt at the
// standard 96 CSS-px-per-inch reference (same reference DOCX's px→twip
// conversion already assumes), matching how a browser would size the row.
const POINTS_PER_PX = 0.75;

// X6 — no font metrics available, so column width is approximated from the
// longest flattened cell text found in that column (across every row),
// clamped to a sane range — the same category of "character count is the
// best available proxy" heuristic already used for DOCX's table columns.
const MIN_COL_WIDTH = 8;
const MAX_COL_WIDTH = 60;

// Excel's column `width` unit is roughly one "0"-glyph's width in the
// workbook's default font (11pt Calibri) — approximated at ~7px (~5.25pt) per
// the same 96px/in reference POINTS_PER_PX uses, per ECMA-376's own column
// width definition. Used only to fold a cell's own left+right CSS padding
// into the width heuristic below — a `<td style="padding:8px">` visibly
// widens its column in a browser, which the plain "longest text" proxy alone
// has no way to account for.
const PT_PER_CHAR_UNIT = 5.25;

export function computeColumnWidths(grid, columnCount) {
  const maxLen = new Array(columnCount).fill(0);
  for (const gridRow of grid) {
    for (const [col, cell] of gridRow) {
      const paddingChars = ((cell.padding?.left ?? 0) + (cell.padding?.right ?? 0)) / PT_PER_CHAR_UNIT;
      maxLen[col] = Math.max(maxLen[col], flattenCellText(cell).length + paddingChars);
    }
  }
  return maxLen.map((len) => Math.min(MAX_COL_WIDTH, Math.max(MIN_COL_WIDTH, len + 2)));
}

// A row with no explicit CSS `height` still visibly grows in a browser to
// fit its tallest cell's own font size, line count, and top+bottom padding
// (`<td style="padding:8px">` is a common, otherwise-invisible-to-Excel,
// source of extra row height) — mirrors document-sheet.js's font-aware
// height for heading/paragraph rows, applied per-cell here since a table
// row's height is governed by whichever cell in it needs the most space.
// Returns undefined when every cell is empty and unpadded (nothing to add
// beyond Excel's own default row height).
const DEFAULT_FONT_SIZE_PT = 11;
// See document-sheet.js's identical constant for the full reasoning: a plain
// 1.2x leading under-estimates real Excel/Google Sheets line height (~1.36x-
// 1.575x for either app's own default single-line row), and that per-line
// shortfall compounds visibly for a multi-line wrapped cell (e.g. a `<br>`-
// joined address in one `<td>`).
const LINE_HEIGHT_MULTIPLIER = 1.5;

function estimateRowHeightPt(gridRow, rctx) {
  let maxHeightPt;
  for (const cell of gridRow.values()) {
    const runs = trimRuns(flattenCellRuns(cell));
    const text = runs.map((r) => r.text).join('');
    const paddingTopPt = cell.padding?.top ?? 0;
    const paddingBottomPt = cell.padding?.bottom ?? 0;
    if (text === '' && paddingTopPt === 0 && paddingBottomPt === 0) continue;
    const fontSizePt = Math.max(rctx?.defaultFontSizePt ?? DEFAULT_FONT_SIZE_PT, ...runs.map((r) => r.fontSizePt ?? 0));
    const lines = text === '' ? 1 : text.split('\n').length;
    const heightPt = lines * fontSizePt * LINE_HEIGHT_MULTIPLIER + paddingTopPt + paddingBottomPt;
    maxHeightPt = maxHeightPt === undefined ? heightPt : Math.max(maxHeightPt, heightPt);
  }
  return maxHeightPt;
}

// `writing-mode:vertical-rl` maps to Excel's dedicated "stacked" vertical
// text value (255) — takes priority since it's a distinct layout mode, not
// just an angle. Otherwise an explicit `transform:rotate(Ndeg)` maps to a
// real rotation angle: CSS's positive (clockwise) degrees become 90+N per
// ST_CellAlignmentTextRotation (Excel's clockwise-angle range is 91-180,
// value = 90 + degrees); negative (counterclockwise) degrees become their
// absolute value (Excel's counterclockwise range is 0-90 directly).
function cellTextRotation(cell) {
  if (cell.writingMode === 'vertical-rl') return 255;
  if (cell.rotateDeg === undefined) return undefined;
  const deg = Math.max(-90, Math.min(90, cell.rotateDeg));
  return deg >= 0 ? 90 + deg : -deg;
}

function renderCell(b, cell, rowIndex, col, rctx) {
  const ref = cellRef(rowIndex, col);
  const runs = trimRuns(flattenCellRuns(cell));
  const text = runs.map((r) => r.text).join('');
  const value = classifyCellValue(text);

  // Font styling (bold/italic/underline/strikethrough/color/size) comes from
  // the single run when there's exactly one — flattenCellRuns already folds
  // in the cell's own style plus whatever single inline wrapper it had, so
  // this is the one true source, not `cell.bold`/`cell.color` read
  // separately (which would silently miss italic/underline/strikethrough
  // entirely, and diverge from what the text itself actually looks like).
  // 0 runs (empty) or >1 (going rich — each run carries its own font) both
  // mean the cell itself has no single font to fall back to.
  const fontStyle = runs.length === 1 ? runs[0] : {};

  const styleIndex = rctx.styles.registerCellStyle({
    ...fontStyle,
    backgroundColor: cell.backgroundColor,
    borderTop: cell.borderTop,
    borderRight: cell.borderRight,
    borderBottom: cell.borderBottom,
    borderLeft: cell.borderLeft,
    numFmtCode: value.numFmtCode,
    align: cell.align,
    // An explicit CSS vertical-align always wins; absent that, Excel's own
    // default is BOTTOM (not middle) — invisible for a short single-line
    // row, but estimateRowHeightPt now often makes a padded row noticeably
    // taller than its text needs, so a bottom-stuck line inside it would
    // look visibly wrong next to the browser's symmetric-padding rendering.
    // Only defaulted for a cell with actual content — a truly empty,
    // unstyled cell must stay on the shared default style (index 0) so it
    // can keep omitting its own <c> entirely (see the value.type === 'empty'
    // branch below).
    verticalAlign: value.type === 'empty' ? cell.verticalAlign : (cell.verticalAlign ?? 'center'),
    readingOrder: cell.direction === 'rtl' ? 2 : undefined,
    textRotation: cellTextRotation(cell),
    // A cell whose content came from several block children (e.g. a <td>
    // holding an <h3> plus multiple <p>s — see flattenCellRuns' own '\n'-
    // joining) or a `<br>`-joined run within one paragraph carries real
    // newline characters in `text` — but a plain (non-rich) Excel string
    // renders every character on ONE visual line, newlines included, unless
    // the cell's own alignment explicitly sets wrapText (see styles.js's
    // own `xf.wrapText` -> `<alignment wrapText="1"/>` comment). Without it,
    // several paragraphs' worth of text visually ran together with no
    // separation at all (estimateRowHeightPt already sizes the ROW tall
    // enough for every line — this is what makes Excel actually break onto
    // them). Matches document-sheet.js's own identical `wrapText` field for
    // a top-level block row.
    wrapText: value.type !== 'number' && text.split('\n').length > 1,
  });

  if (value.type === 'empty') {
    // A styled-but-empty cell (a colored spacer, a bordered blank header
    // slot) still needs its own <c> so the fill/border actually renders —
    // only a truly plain empty cell (default style, index 0) is omitted.
    if (styleIndex !== 0) b.self('c', { r: ref, s: styleIndex });
    return;
  }

  if (value.type === 'number') {
    b.el('c', styleIndex !== 0 ? { r: ref, s: styleIndex } : { r: ref });
    b.el('v').text(value.value).up();
    b.up();
    return;
  }

  // Rich text (multiple differently-styled runs) only when there's genuine
  // per-run style variety left after merging adjacent identical runs — the
  // overwhelmingly common uniformly-styled/plain cell stays on the leaner
  // plain shared-string path, exactly as before.
  const stringIndex = runs.length > 1 ? rctx.sharedStrings.addRich(runs) : rctx.sharedStrings.add(text);
  b.el('c', styleIndex !== 0 ? { r: ref, t: 's', s: styleIndex } : { r: ref, t: 's' });
  b.el('v').text(String(stringIndex)).up();
  b.up();
}

// A merged range (colspan and/or rowspan > 1) only ever gets ONE real <c>
// written — the origin's — since SpreadsheetML treats a missing cell
// reference as blank and a <mergeCell> alone tells Excel to visually span
// that range (see layoutGrid's own doc comment). But Excel derives a merged
// range's BORDER on each edge from whichever underlying cell sits at that
// exact row/column, not from the origin's style stretched across the whole
// rectangle — so the origin's own borderRight/borderBottom (declared for
// its own single cell) only ever rendered at the origin's own position,
// nowhere near the merge's actual right/bottom edge, silently leaving those
// edges with no visible border at all (e.g. this project's own `<td
// colspan="4">Subtotal</td>` — real bordered content in the source HTML,
// but only its own leftmost column ever carried that border in the output).
// One placeholder <c> per OTHER covered position fixes this: top/bottom
// border only on the merge's own top/bottom row, left/right only on its own
// left/right column (an internal position gets none of either, matching a
// real merged cell's single unbroken interior), with the origin's own fill
// repeated too — Excel itself usually auto-extends a merge's fill from the
// origin alone, but not every OOXML consumer does, and it's free to set.
function buildBorderPlaceholders(merges) {
  const byRow = new Map(); // rowIndex -> [{ col, backgroundColor, borderTop, borderRight, borderBottom, borderLeft }]
  for (const { r1, c1, r2, c2, cell } of merges) {
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        if (r === r1 && c === c1) continue; // the origin — already a real <c>
        const placeholder = {
          col: c,
          backgroundColor: cell.backgroundColor,
          borderTop: r === r1 ? cell.borderTop : undefined,
          borderBottom: r === r2 ? cell.borderBottom : undefined,
          borderLeft: c === c1 ? cell.borderLeft : undefined,
          borderRight: c === c2 ? cell.borderRight : undefined,
        };
        if (!byRow.has(r)) byRow.set(r, []);
        byRow.get(r).push(placeholder);
      }
    }
  }
  return byRow;
}

function renderBorderPlaceholderCell(b, placeholder, rowIndex, col, rctx) {
  const styleIndex = rctx.styles.registerCellStyle({
    backgroundColor: placeholder.backgroundColor,
    borderTop: placeholder.borderTop,
    borderRight: placeholder.borderRight,
    borderBottom: placeholder.borderBottom,
    borderLeft: placeholder.borderLeft,
  });
  if (styleIndex !== 0) b.self('c', { r: cellRef(rowIndex, col), s: styleIndex });
}

/**
 * Renders one table's grid as structured, position-independent data — every
 * row/col/merge coordinate already absolute (offset by `startRow`/`startCol`)
 * — rather than a standalone `<worksheet>` document. This is what lets the
 * whole-document sheet builder (document-sheet.js) place several tables, plus
 * heading/paragraph rows, into one shared `<sheetData>` at whatever row/column
 * each block happens to start at. Each row is pre-serialized to an XML string
 * (via a scratch XmlBuilder) so the caller can just concatenate rows from
 * different sources in row order without needing to share a single open
 * builder across them.
 */
export function renderTableGrid(table, rctx, { startRow = 0, startCol = 0 } = {}) {
  const headRows = table.head ?? [];
  const allRows = expandRowsForBands([...headRows, ...table.rows]);
  const { grid, merges, heightsPx, columnCount } = layoutGrid(allRows);
  const placeholdersByRow = buildBorderPlaceholders(merges);

  const rows = grid.map((gridRow, rowIndex) => {
    const heightPx = heightsPx[rowIndex];
    const explicitHeightPt = heightPx != null ? heightPx * POINTS_PER_PX : undefined;
    // An explicit CSS height is a floor, not a cap — same as a browser table
    // row, which still grows to fit content taller than a declared height.
    const computedHeightPt = estimateRowHeightPt(gridRow, rctx);
    const heightPt = explicitHeightPt != null
      ? (computedHeightPt != null ? Math.max(explicitHeightPt, computedHeightPt) : explicitHeightPt)
      : computedHeightPt;

    const rowAttrs = { r: rowIndex + startRow + 1 };
    if (heightPt != null) {
      rowAttrs.ht = heightPt;
      rowAttrs.customHeight = 1;
    }
    const b = new XmlBuilder();
    b.el('row', rowAttrs);
    const placeholders = placeholdersByRow.get(rowIndex) ?? [];
    const cols = [...new Set([...gridRow.keys(), ...placeholders.map((p) => p.col)])].sort((a, b2) => a - b2);
    for (const col of cols) {
      if (gridRow.has(col)) {
        renderCell(b, gridRow.get(col), rowIndex + startRow, col + startCol, rctx);
      } else {
        renderBorderPlaceholderCell(b, placeholders.find((p) => p.col === col), rowIndex + startRow, col + startCol, rctx);
      }
    }
    b.up();
    return { rowIndex: rowIndex + startRow, xml: b.end() };
  });

  const merges2 = merges.map((m) => ({
    r1: m.r1 + startRow, c1: m.c1 + startCol,
    r2: m.r2 + startRow, c2: m.c2 + startCol,
  }));

  return { rows, merges: merges2, columnCount, colWidths: computeColumnWidths(grid, columnCount) };
}

export function worksheetXml(table, rctx) {
  const { rows, merges, colWidths } = renderTableGrid(table, rctx, { startRow: 0, startCol: 0 });

  const b = new XmlBuilder();
  b.el('worksheet', { xmlns: NS_MAIN });

  if (colWidths.length) {
    b.el('cols');
    colWidths.forEach((width, i) => b.self('col', { min: i + 1, max: i + 1, width, customWidth: 1 }));
    b.up();
  }

  b.el('sheetData');
  for (const row of rows) b.raw(row.xml);
  b.up();

  if (merges.length > 0) {
    b.el('mergeCells', { count: merges.length });
    for (const m of merges) b.self('mergeCell', { ref: `${cellRef(m.r1, m.c1)}:${cellRef(m.r2, m.c2)}` });
    b.up();
  }

  return xmlDoc(b.end());
}
