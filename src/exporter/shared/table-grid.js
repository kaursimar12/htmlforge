/**
 * Expands a TableNode's rows (each row only listing the cells that actually
 * start there, exactly like HTML) into a full row×column grid — every
 * rowspan/colspan's covered position gets a real placeholder slot pointing
 * back at the cell that covers it (`originCell`), rather than being left
 * absent. Shared by every exporter whose target format needs a real
 * rectangular grid to walk (OOXML tables need every row to carry the same
 * column count; a `<w:tc>`/`<a:tc>` continuation slot still needs to render
 * its own border/shading — see docx/document.js's vMerge comment for why
 * `originCell` matters, not just a bare "skip this cell" flag).
 */
export function expandTableGrid(rows) {
  const pending = new Map(); // column index -> { rowsLeft, colspan, cell }
  const grid = [];
  let maxCols = 0;

  for (const row of rows) {
    const gridRow = [];
    let col = 0;
    let cellIdx = 0;

    while (cellIdx < row.cells.length || [...pending.keys()].some((c) => c >= col)) {
      const p = pending.get(col);
      if (p && p.rowsLeft > 0) {
        gridRow.push({ cell: null, colspan: p.colspan, continuation: true, originCell: p.cell });
        p.rowsLeft--;
        if (p.rowsLeft === 0) pending.delete(col);
        col += p.colspan;
        continue;
      }
      if (cellIdx >= row.cells.length) {
        const next = [...pending.keys()].filter((c) => c >= col).sort((a, b) => a - b)[0];
        if (next === undefined) break;
        col = next;
        continue;
      }
      const cell = row.cells[cellIdx++];
      const colspan = cell.colspan ?? 1;
      const rowspan = cell.rowspan ?? 1;
      gridRow.push({ cell, colspan, restart: rowspan > 1 });
      if (rowspan > 1) pending.set(col, { rowsLeft: rowspan - 1, colspan, cell });
      col += colspan;
    }

    // gridRow is an array of cell slots, but it also carries the source
    // row's own heightPx (if declared) as a plain property — simplest way to
    // keep it attached to the row without changing every consumer's shape.
    gridRow.heightPx = row.heightPx;
    grid.push(gridRow);
    maxCols = Math.max(maxCols, col);
  }

  return { grid, columnCount: maxCols || 1 };
}
