/**
 * CSV exporter — IDM `TableNode`s → RFC 4180 CSV.
 * Milestones: C1 (first table → rows) · C2 (RFC 4180 quoting) ·
 *             C3 (multiple tables → concatenated, blank line between)
 *
 * Each table is labeled with its own `<caption>` if present, otherwise the
 * nearest preceding heading in document order (if any) — rendered as a
 * one-line label row above that table's data rows.
 */

import { flattenCellText, tableRows } from '../../idm/flatten-text.js';
import { collectTables } from '../../idm/collect-tables.js';

export const csvExporter = {
  id: 'csv',
  extensions: ['.csv'],
  mimeType: 'text/csv',

  export(document, ctx) {
    const found = [];
    collectTables(document.children, { lastHeading: undefined }, found);

    if (found.length === 0) {
      ctx.warn('no-table', 'No <table> found in the document; CSV output is empty');
      return Buffer.from('');
    }

    if (found.length > 1) {
      ctx.warn(
        'multiple-tables-concatenated',
        `Found ${found.length} tables; concatenated into one CSV, separated by a blank line`,
      );
    }

    const csv = found.map(({ table, label }) => renderTable(table, label)).join('\r\n\r\n');
    return Buffer.from(csv + '\r\n', 'utf8');
  },
};

// ─── C1 — rows ────────────────────────────────────────────────────────────────

function renderTable(table, label) {
  const rowsCsv = tableRows(table).map(renderRow).join('\r\n');

  if (!label) return rowsCsv;

  return quoteField(label) + '\r\n' + rowsCsv;
}

function renderRow(row) {
  return row.cells.map((cell) => quoteField(flattenCellText(cell))).join(',');
}

// ─── C2 — RFC 4180 quoting ────────────────────────────────────────────────────

function quoteField(field) {
  if (/[",\r\n]/.test(field)) {
    return '"' + field.replace(/"/g, '""') + '"';
  }
  return field;
}
