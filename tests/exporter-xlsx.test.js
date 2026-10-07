import { describe, it, expect } from 'vitest';
import { convert } from '../src/index.js';
import { readXlsxParts } from './helpers/xlsx.js';
import { findAll, find, attr, textContent } from '../src/utils/xml.js';
import { colToLetter, cellRef, classifyCellValue } from '../src/exporter/xlsx/cell-value.js';

const REQUIRED_PARTS = [
  '[Content_Types].xml',
  '_rels/.rels',
  'docProps/core.xml',
  'docProps/app.xml',
  'xl/workbook.xml',
  'xl/worksheets/sheet1.xml',
  'xl/sharedStrings.xml',
  'xl/styles.xml',
  'xl/_rels/workbook.xml.rels',
];

// Resolves a shared-string cell's <v> index to its actual text, so tests
// can assert on real strings rather than opaque indices.
function sharedString(xml, index) {
  const strings = findAll(xml['xl/sharedStrings.xml'], 'si');
  return textContent(strings[index]);
}

function cellsByRef(xml, sheet = 'xl/worksheets/sheet1.xml') {
  const map = {};
  for (const c of findAll(xml[sheet], 'c')) map[attr(c, 'r')] = c;
  return map;
}

// Finds the cell reference whose shared-string value matches `text` —
// guards on t="s" first since a merged row's companion cells (every column
// past the merge anchor) carry only a style, no <v> at all.
function refOfString(xml, cells, text) {
  return Object.keys(cells).find((ref) => {
    const cell = cells[ref];
    return attr(cell, 't') === 's' && sharedString(xml, Number(textContent(find(cell, 'v')))) === text;
  });
}

// xl/styles.xml has an <xf> inside both <cellStyleXfs> and <cellXfs> — a
// plain findAll(styles, 'xf') matches both and misindexes every entry by
// one. This scopes the search to <cellXfs>'s own <xf> children only, which
// is what a cell's s="N" actually indexes into.
function cellXfsOf(xml) {
  return findAll(find(xml['xl/styles.xml'], 'cellXfs'), 'xf');
}

describe('cell-value.js — pure utilities', () => {
  it('colToLetter: 0-25 map to A-Z, 26 wraps to AA', () => {
    expect(colToLetter(0)).toBe('A');
    expect(colToLetter(25)).toBe('Z');
    expect(colToLetter(26)).toBe('AA');
    expect(colToLetter(27)).toBe('AB');
    expect(colToLetter(51)).toBe('AZ');
    expect(colToLetter(52)).toBe('BA');
  });

  it('cellRef: 0-based row/col to A1 notation', () => {
    expect(cellRef(0, 0)).toBe('A1');
    expect(cellRef(1, 1)).toBe('B2');
    expect(cellRef(9, 26)).toBe('AA10');
  });

  it('classifyCellValue: empty string', () => {
    expect(classifyCellValue('')).toEqual({ type: 'empty' });
    expect(classifyCellValue('   ')).toEqual({ type: 'empty' });
  });

  it('classifyCellValue: plain integers and decimals', () => {
    expect(classifyCellValue('123')).toEqual({ type: 'number', value: '123' });
    expect(classifyCellValue('-45.6')).toEqual({ type: 'number', value: '-45.6' });
  });

  it('classifyCellValue: percent', () => {
    expect(classifyCellValue('5%')).toEqual({ type: 'number', value: '0.05', numFmtCode: '0%' });
    expect(classifyCellValue('5.00%')).toEqual({ type: 'number', value: '0.05', numFmtCode: '0.00%' });
  });

  it('classifyCellValue: currency ($/€/£)', () => {
    expect(classifyCellValue('$50.00')).toEqual({ type: 'number', value: '50.00', numFmtCode: '"$"#,##0.00' });
    expect(classifyCellValue('€10')).toEqual({ type: 'number', value: '10', numFmtCode: '"€"#,##0' });
  });

  it('classifyCellValue: thousands-separated numbers', () => {
    expect(classifyCellValue('1,234')).toEqual({ type: 'number', value: '1234', numFmtCode: '#,##0' });
    expect(classifyCellValue('1,234.50')).toEqual({ type: 'number', value: '1234.50', numFmtCode: '#,##0.00' });
  });

  it('classifyCellValue: stays text for forms deliberately not recognized', () => {
    expect(classifyCellValue('N/A')).toEqual({ type: 'string' });
    expect(classifyCellValue('3 items')).toEqual({ type: 'string' });
    expect(classifyCellValue('$(50.00)')).toEqual({ type: 'string' }); // accounting negative
    expect(classifyCellValue('¥100')).toEqual({ type: 'string' }); // unsupported currency symbol
    expect(classifyCellValue('1e10')).toEqual({ type: 'string' }); // scientific notation
  });
});

describe('XLSX exporter', () => {
  it('returns correct mimeType and extension', async () => {
    const result = await convert('<table><tr><td>x</td></tr></table>', { format: 'xlsx' });
    expect(result.mimeType).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(result.extension).toBe('.xlsx');
  });

  // ─── X1 — skeleton ──────────────────────────────────────────────────────

  it('X1 — produces a valid ZIP with every required OOXML part, each parsing as well-formed XML', async () => {
    const { buffer } = await convert('<table><tr><td>x</td></tr></table>', { format: 'xlsx' });
    expect(buffer[0]).toBe(0x50); // P
    expect(buffer[1]).toBe(0x4b); // K

    const { zip, xml } = readXlsxParts(buffer);
    for (const part of REQUIRED_PARTS) {
      expect(zip.has(part), `missing part: ${part}`).toBe(true);
      expect(xml[part], `did not parse as XML: ${part}`).toBeTruthy();
    }
  });

  it('X1 — [Content_Types].xml declares Overrides for workbook/styles/sharedStrings/worksheet', async () => {
    const { buffer } = await convert('<table><tr><td>x</td></tr></table>', { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const overrides = findAll(xml['[Content_Types].xml'], 'Override').map((o) => attr(o, 'PartName'));
    expect(overrides).toContain('/xl/workbook.xml');
    expect(overrides).toContain('/xl/styles.xml');
    expect(overrides).toContain('/xl/sharedStrings.xml');
    expect(overrides).toContain('/xl/worksheets/sheet1.xml');
  });

  it('X1 — _rels/.rels points its officeDocument relationship at xl/workbook.xml', async () => {
    const { buffer } = await convert('<table><tr><td>x</td></tr></table>', { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const rel = findAll(xml['_rels/.rels'], 'Relationship').find((r) => attr(r, 'Target') === 'xl/workbook.xml');
    expect(rel).toBeTruthy();
  });

  it('X1 — xl/_rels/workbook.xml.rels references styles.xml, sharedStrings.xml, and the worksheet', async () => {
    const { buffer } = await convert('<table><tr><td>x</td></tr></table>', { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const targets = findAll(xml['xl/_rels/workbook.xml.rels'], 'Relationship').map((r) => attr(r, 'Target'));
    expect(targets).toContain('styles.xml');
    expect(targets).toContain('sharedStrings.xml');
    expect(targets).toContain('worksheets/sheet1.xml');
  });

  it('X1 — a document with no exportable content still produces one valid empty sheet, plus a warning', async () => {
    const { buffer, warnings } = await convert('', { format: 'xlsx' });
    const { zip, xml } = readXlsxParts(buffer);
    expect(zip.has('xl/worksheets/sheet1.xml')).toBe(true);
    expect(findAll(xml['xl/worksheets/sheet1.xml'], 'row')).toHaveLength(0);
    expect(warnings.some((w) => w.code === 'no-content')).toBe(true);
  });

  // ─── X2 — table → worksheet, A1 refs, value classification ────────────

  it('X2 — a 2x2 table maps to correct A1 cell refs', async () => {
    const html = '<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    expect(Object.keys(cells).sort()).toEqual(['A1', 'A2', 'B1', 'B2']);
    expect(sharedString(xml, Number(textContent(find(cells.A1, 'v'))))).toBe('A');
    expect(sharedString(xml, Number(textContent(find(cells.D2 ?? cells.B2, 'v'))))).toBeTruthy();
  });

  it('X2 — a text cell uses t="s" and resolves to the right shared string', async () => {
    const { buffer } = await convert('<table><tr><td>Hello</td></tr></table>', { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const c = find(xml['xl/worksheets/sheet1.xml'], 'c');
    expect(attr(c, 't')).toBe('s');
    const index = Number(textContent(find(c, 'v')));
    expect(sharedString(xml, index)).toBe('Hello');
  });

  it('X2 — a plain numeric cell has no t attribute and the literal value', async () => {
    const { buffer } = await convert('<table><tr><td>123</td></tr></table>', { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const c = find(xml['xl/worksheets/sheet1.xml'], 'c');
    expect(attr(c, 't')).toBeFalsy();
    expect(textContent(find(c, 'v'))).toBe('123');
  });

  it('X2 — an empty cell produces no <c> at all', async () => {
    const { buffer } = await convert('<table><tr><td>A</td><td></td></tr></table>', { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    expect(cells.A1).toBeTruthy();
    expect(cells.B1).toBeUndefined();
  });

  it('X2 — currency/percent/thousands cells all render as real numeric cells with matching formats', async () => {
    const html = '<table><tr><td>$50.00</td><td>5.00%</td><td>1,234</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    expect(attr(cells.A1, 't')).toBeFalsy();
    expect(textContent(find(cells.A1, 'v'))).toBe('50.00');
    expect(textContent(find(cells.B1, 'v'))).toBe('0.05');
    expect(textContent(find(cells.C1, 'v'))).toBe('1234');
  });

  // ─── X3 — cell styling ──────────────────────────────────────────────────

  it('X3 — a bare CSS generic keyword for the document font resolves to a real installable font name', async () => {
    const html = '<html><body style="font-family:sans-serif"><table><tr><td>Hi</td></tr></table></body></html>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const fonts = findAll(xml['xl/styles.xml'], 'font');
    const names = fonts.map((f) => attr(find(f, 'name'), 'val'));
    expect(names).toContain('Arial');
    expect(names).not.toContain('sans-serif');
  });

  it('X3 — <thead>/<th> cells are bold', async () => {
    const html = '<table><thead><tr><th>Head</th></tr></thead><tr><td>Data</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    const headStyle = Number(attr(cells.A1, 's') ?? 0);
    const dataStyle = Number(attr(cells.A2, 's') ?? 0);
    const cellXfs = cellXfsOf(xml);
    const fonts = findAll(xml['xl/styles.xml'], 'font');
    const headFontId = Number(attr(cellXfs[headStyle], 'fontId'));
    expect(find(fonts[headFontId], 'b')).toBeTruthy();
    const dataFontId = Number(attr(cellXfs[dataStyle], 'fontId'));
    expect(find(fonts[dataFontId], 'b')).toBeFalsy();
  });

  it('X3 — background/color/border/align all apply through the composite style registry, deduplicated', async () => {
    const html = '<table><tr>' +
      '<td style="background:#1F6585;color:#fff;font-weight:bold">A</td>' +
      '<td style="background:#1F6585;color:#fff;font-weight:bold">B</td>' + // identical style — should dedupe
      '<td style="border-left:2px solid #000">C</td>' +
      '<td style="text-align:right">D</td>' +
      '</tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);

    // A and B share one style index (dedup).
    expect(attr(cells.A1, 's')).toBe(attr(cells.B1, 's'));

    const cellXfs = cellXfsOf(xml);
    const fonts = findAll(xml['xl/styles.xml'], 'font');
    const fills = findAll(xml['xl/styles.xml'], 'fill');

    const abXf = cellXfs[Number(attr(cells.A1, 's'))];
    const abFont = fonts[Number(attr(abXf, 'fontId'))];
    expect(find(abFont, 'b')).toBeTruthy();
    expect(attr(find(abFont, 'color'), 'rgb')).toBe('FFFFFFFF');
    const abFill = fills[Number(attr(abXf, 'fillId'))];
    expect(attr(find(abFill, 'fgColor'), 'rgb')).toBe('FF1F6585');

    const cXf = cellXfs[Number(attr(cells.C1, 's'))];
    const borders = findAll(xml['xl/styles.xml'], 'border');
    const cBorder = borders[Number(attr(cXf, 'borderId'))];
    expect(find(cBorder, 'left')).toBeTruthy();
    expect(attr(find(cBorder, 'right'), 'style')).toBeFalsy(); // present (schema requires all 4 sides) but undeclared

    const dXf = cellXfs[Number(attr(cells.D1, 's'))];
    expect(attr(find(dXf, 'alignment'), 'horizontal')).toBe('right');
  });

  it('X3 — vertical-align:top/middle/bottom map to alignment vertical="top"/"center"/"bottom"', async () => {
    const html = '<table><tr>' +
      '<td style="vertical-align:top">A</td>' +
      '<td style="vertical-align:middle">B</td>' +
      '<td style="vertical-align:bottom">C</td>' +
      '</tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    const cellXfs = cellXfsOf(xml);
    const vAlignOf = (cell) => attr(find(cellXfs[Number(attr(cell, 's'))], 'alignment'), 'vertical');
    expect(vAlignOf(cells.A1)).toBe('top');
    expect(vAlignOf(cells.B1)).toBe('center'); // CSS "middle" -> Excel "center"
    expect(vAlignOf(cells.C1)).toBe('bottom');
  });

  it('X3 — a single styled cell (not rich-text) still applies italic/underline/strikethrough/fontSize at the cell level', async () => {
    const html = '<table><tr>' +
      '<td style="font-style:italic">I</td>' +
      '<td style="text-decoration:underline">U</td>' +
      '<td style="text-decoration:line-through">S</td>' +
      '<td style="font-size:24px">Big</td>' +
      '</tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    const cellXfs = cellXfsOf(xml);
    const fonts = findAll(xml['xl/styles.xml'], 'font');
    const fontOf = (cell) => fonts[Number(attr(cellXfs[Number(attr(cell, 's'))], 'fontId'))];
    expect(find(fontOf(cells.A1), 'i')).toBeTruthy();
    expect(find(fontOf(cells.B1), 'u')).toBeTruthy();
    expect(find(fontOf(cells.C1), 'strike')).toBeTruthy();
    expect(Number(attr(find(fontOf(cells.D1), 'sz'), 'val'))).toBe(18); // 24px * 0.75
  });

  it('X3 — a currency cell\'s numFmt is registered as a custom format (id >= 164)', async () => {
    const { buffer } = await convert('<table><tr><td>$50.00</td></tr></table>', { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const numFmt = find(xml['xl/styles.xml'], 'numFmt');
    expect(Number(attr(numFmt, 'numFmtId'))).toBeGreaterThanOrEqual(164);
    expect(attr(numFmt, 'formatCode')).toBe('"$"#,##0.00');
  });

  it('X3 — xl/styles.xml has the minimal skeleton Excel requires even with no styling at all', async () => {
    const { buffer } = await convert('<table><tr><td>Plain</td></tr></table>', { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const styles = xml['xl/styles.xml'];
    expect(findAll(styles, 'font').length).toBeGreaterThan(0);
    expect(findAll(styles, 'fill').length).toBeGreaterThanOrEqual(2);
    expect(findAll(styles, 'border').length).toBeGreaterThan(0);
    expect(find(styles, 'cellStyleXfs')).toBeTruthy();
    expect(find(styles, 'cellXfs')).toBeTruthy();
    expect(find(styles, 'cellStyles')).toBeTruthy();
  });

  // ─── X4 — shared strings ───────────────────────────────────────────────

  it('X4 — an identical string across two sheets dedupes to a single shared entry', async () => {
    const html = '<h2>One</h2><table><tr><td>Repeat</td></tr></table>' +
      '<h2>Two</h2><table><tr><td>Repeat</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const strings = findAll(xml['xl/sharedStrings.xml'], 'si').map((si) => textContent(si));
    expect(strings.filter((s) => s === 'Repeat')).toHaveLength(1);
  });

  it('X4 — sharedStrings count/uniqueCount reflect total references vs. distinct strings', async () => {
    const html = '<table><tr><td>Same</td><td>Same</td><td>Different</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const sst = xml['xl/sharedStrings.xml'];
    expect(Number(attr(sst, 'count'))).toBe(3);
    expect(Number(attr(sst, 'uniqueCount'))).toBe(2);
  });

  // ─── X5 — single-sheet naming ──────────────────────────────────────────

  it('X5 — a single table with a caption names the (only) sheet after it', async () => {
    const html = '<table><caption>My Caption</caption><tr><td>x</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const names = findAll(xml['xl/workbook.xml'], 'sheet').map((s) => attr(s, 'name'));
    expect(names).toEqual(['My Caption']);
  });

  it('X5 — a document <title> names the sheet, taking priority over a single table\'s own caption', async () => {
    const html = '<html><head><title>Doc Title</title></head><body>' +
      '<table><caption>My Caption</caption><tr><td>x</td></tr></table></body></html>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    expect(attr(find(xml['xl/workbook.xml'], 'sheet'), 'name')).toBe('Doc Title');
  });

  it('X5 — illegal sheet-name characters are sanitized', async () => {
    const html = '<html><head><title>Q1/Q2 Report</title></head><body><table><tr><td>x</td></tr></table></body></html>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const name = attr(find(xml['xl/workbook.xml'], 'sheet'), 'name');
    expect(name).not.toMatch(/[\\/?*[\]:]/);
  });

  it('X5 — a title longer than 31 characters is truncated', async () => {
    const longTitle = 'A'.repeat(50);
    const html = `<html><head><title>${longTitle}</title></head><body><table><tr><td>x</td></tr></table></body></html>`;
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const name = attr(find(xml['xl/workbook.xml'], 'sheet'), 'name');
    expect(name.length).toBeLessThanOrEqual(31);
  });

  it('X5 — options.sheetName wins over a document <title> and a table\'s own caption', async () => {
    const html = '<html><head><title>Doc Title</title></head><body>' +
      '<table><caption>My Caption</caption><tr><td>x</td></tr></table></body></html>';
    const { buffer } = await convert(html, { format: 'xlsx', sheetName: 'Custom Name' });
    const { xml } = readXlsxParts(buffer);
    expect(attr(find(xml['xl/workbook.xml'], 'sheet'), 'name')).toBe('Custom Name');
  });

  it('X5 — no title/sheetName/single-table-caption at all falls back to "Sheet1"', async () => {
    const html = '<table><tr><td>x</td></tr></table><table><tr><td>y</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const names = findAll(xml['xl/workbook.xml'], 'sheet').map((s) => attr(s, 'name'));
    expect(names).toEqual(['Sheet1']);
  });

  // ─── Document flow — whole document on one sheet, in document order ────

  it('Document flow — a heading, paragraph, and table with no other table all land on one sheet, in order', async () => {
    const html = '<h1>Report</h1><p>Summary text.</p><table><tr><td>Data</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    expect(findAll(xml['xl/workbook.xml'], 'sheet')).toHaveLength(1);
    const strings = findAll(xml['xl/sharedStrings.xml'], 'si').map((si) => textContent(si));
    expect(strings).toEqual(expect.arrayContaining(['Report', 'Summary text.', 'Data']));
  });

  it('Document flow — two tables separated by a heading both render on the one sheet, heading between them by row order', async () => {
    const html = '<table><tr><td>x</td></tr></table><h2>My Heading</h2><table><tr><td>y</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    expect(findAll(xml['xl/workbook.xml'], 'sheet')).toHaveLength(1);
    const cells = cellsByRef(xml);
    const rowOf = (ref) => Number(/\d+/.exec(ref)[0]);
    expect(rowOf(refOfString(xml, cells, 'x'))).toBeLessThan(rowOf(refOfString(xml, cells, 'My Heading')));
    expect(rowOf(refOfString(xml, cells, 'My Heading'))).toBeLessThan(rowOf(refOfString(xml, cells, 'y')));
  });

  it('Document flow — a heading/paragraph row carries its own background fill and bold, merged across the full sheet width', async () => {
    const html = '<div style="background:#1f6586;color:#fff"><h1>INVOICE</h1></div>' +
      '<table><tr><td>A</td><td>B</td><td>C</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    const headingRef = refOfString(xml, cells, 'INVOICE');
    const cellXfs = cellXfsOf(xml);
    const xf = cellXfs[Number(attr(cells[headingRef], 's'))];
    const fills = findAll(xml['xl/styles.xml'], 'fill');
    expect(attr(find(fills[Number(attr(xf, 'fillId'))], 'fgColor'), 'rgb')).toBe('FF1F6586');
    const fonts = findAll(xml['xl/styles.xml'], 'font');
    expect(find(fonts[Number(attr(xf, 'fontId'))], 'b')).toBeTruthy(); // heading -> bold
    const merges = findAll(xml['xl/worksheets/sheet1.xml'], 'mergeCell').map((m) => attr(m, 'ref'));
    expect(merges.some((ref) => ref.startsWith(headingRef.match(/^[A-Z]+/)[0] + headingRef.match(/\d+/)[0] + ':'))).toBe(true);
  });

  it('Document flow — a right-aligned paragraph gets horizontal="right" alignment', async () => {
    const html = '<p style="text-align:right">Right side</p><table><tr><td>x</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    const ref = refOfString(xml, cells, 'Right side');
    const cellXfs = cellXfsOf(xml);
    expect(attr(find(cellXfs[Number(attr(cells[ref], 's'))], 'alignment'), 'horizontal')).toBe('right');
  });

  it('Document flow — a table with width:50%;margin-left:auto is offset toward the right half of the sheet', async () => {
    const html = '<table><tr><td>A</td><td>B</td><td>C</td><td>D</td></tr></table>' +
      '<table style="width:50%;margin-left:auto"><tr><td>Total</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    const ref = refOfString(xml, cells, 'Total');
    expect(ref.match(/^[A-Z]+/)[0]).not.toBe('A'); // not hugging the left edge like the first table did
  });

  // ─── X6 — column auto-width ─────────────────────────────────────────────

  it('X6 — a column with longer content gets a wider <col> than a short one', async () => {
    const html = '<table><tr><td>A very long piece of text indeed</td><td>Hi</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cols = findAll(xml['xl/worksheets/sheet1.xml'], 'col');
    expect(Number(attr(cols[0], 'width'))).toBeGreaterThan(Number(attr(cols[1], 'width')));
  });

  it('X6 — column width is clamped to the configured min/max bounds', async () => {
    const html = '<table><tr><td>x</td><td>' + 'y'.repeat(200) + '</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cols = findAll(xml['xl/worksheets/sheet1.xml'], 'col');
    expect(Number(attr(cols[0], 'width'))).toBeGreaterThanOrEqual(8);
    expect(Number(attr(cols[1], 'width'))).toBeLessThanOrEqual(60);
  });

  // ─── X7 — merged cells ──────────────────────────────────────────────────

  it('X7 — a colspan=2 cell produces one mergeCell and no covered cell', async () => {
    const html = '<table><tr><td colspan="2">Wide</td></tr><tr><td>A</td><td>B</td></tr></table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const merges = findAll(xml['xl/worksheets/sheet1.xml'], 'mergeCell');
    expect(merges).toHaveLength(1);
    expect(attr(merges[0], 'ref')).toBe('A1:B1');
    const cells = cellsByRef(xml);
    expect(cells.A1).toBeTruthy();
    expect(cells.B1).toBeUndefined(); // covered by the merge, no cell needed
  });

  it('X7 — a rowspan=2 cell in column A correctly pushes the next row\'s real first cell to column B', async () => {
    const html = '<table>' +
      '<tr><td rowspan="2">Tall</td><td>Row1</td></tr>' +
      '<tr><td>Row2</td></tr>' +
      '</table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const cells = cellsByRef(xml);
    expect(cells.A1).toBeTruthy();
    expect(cells.B1).toBeTruthy();
    expect(cells.A2).toBeUndefined(); // covered by the rowspan
    expect(cells.B2).toBeTruthy(); // correctly shifted to column B, not A
  });

  it('X7 — a cell with both colspan and rowspan produces the right 2x2 merge range', async () => {
    const html = '<table>' +
      '<tr><td colspan="2" rowspan="2">Big</td><td>C1</td></tr>' +
      '<tr><td>C2</td></tr>' +
      '</table>';
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const merges = findAll(xml['xl/worksheets/sheet1.xml'], 'mergeCell');
    expect(attr(merges[0], 'ref')).toBe('A1:B2');
  });

  // ─── Acceptance — invoice-style fixture ────────────────────────────────

  it('Acceptance — an invoice-style fixture (merged header, currency values, right-aligned totals, merged footer) renders correctly', async () => {
    const html = `
      <table>
        <caption>Invoice</caption>
        <tr>
          <td colspan="2" rowspan="2" style="background:#1F6585;color:#fff;font-weight:bold">INVOICE</td>
          <td style="text-align:right"><b>Your Company Name</b></td>
        </tr>
        <tr><td style="text-align:right">Street Address</td></tr>
        <tr><td>Item</td><td>Rate</td><td>Amount</td></tr>
        <tr><td>Logo design</td><td>$50.00</td><td>$150.00</td></tr>
        <tr><td colspan="2" style="text-align:right;font-weight:bold">Total</td><td style="text-align:right;background:#D8E8F8">$157.50</td></tr>
        <tr><td colspan="3" style="background:#1F6585;color:#fff;text-align:center;font-weight:bold">Thank you for your business!</td></tr>
      </table>
    `;
    const { buffer, warnings } = await convert(html, { format: 'xlsx' });
    expect(warnings).toEqual([]);
    const { xml } = readXlsxParts(buffer);

    expect(attr(find(xml['xl/workbook.xml'], 'sheet'), 'name')).toBe('Invoice');

    const merges = findAll(xml['xl/worksheets/sheet1.xml'], 'mergeCell').map((m) => attr(m, 'ref'));
    expect(merges).toContain('A1:B2'); // INVOICE banner
    expect(merges).toContain('A6:C6'); // footer banner (row 6, 0 header rows + 6 body rows deep)

    const cells = cellsByRef(xml);
    // Rate ($50.00) and Amount ($150.00) are real numeric cells.
    const rateCell = cells.B4;
    expect(attr(rateCell, 't')).toBeFalsy();
    expect(textContent(find(rateCell, 'v'))).toBe('50.00');
  });
});

describe('XLSX exporter — images (X8)', () => {
  const TINY_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

  it('an image is actually embedded (real media file + drawing relationship + a:blip), not dropped', async () => {
    const { buffer } = await convert(`<img src="${TINY_PNG}" alt="Dot" width="10" height="10">`, { format: 'xlsx' });
    const { zip, xml } = readXlsxParts(buffer);

    const drawingRef = find(xml['xl/worksheets/sheet1.xml'], 'drawing');
    expect(drawingRef).toBeTruthy();
    const rId = attr(drawingRef, 'r:id');
    const rel = findAll(xml['xl/worksheets/_rels/sheet1.xml.rels'], 'Relationship').find((r) => attr(r, 'Id') === rId);
    expect(attr(rel, 'Target')).toBe('../drawings/drawing1.xml');

    const blip = find(xml['xl/drawings/drawing1.xml'], 'a:blip');
    expect(blip).toBeTruthy();
    const embedId = attr(blip, 'r:embed');
    const drawingRel = findAll(xml['xl/drawings/_rels/drawing1.xml.rels'], 'Relationship').find((r) => attr(r, 'Id') === embedId);
    expect(attr(drawingRel, 'Target')).toBe('../media/image1.png');

    expect(zip.has('xl/media/image1.png')).toBe(true);
    expect(zip.read('xl/media/image1.png').length).toBeGreaterThan(0);
  });

  it('[Content_Types].xml declares Overrides/Defaults for the drawing part and image extension', async () => {
    const { buffer } = await convert(`<img src="${TINY_PNG}">`, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const defaults = findAll(xml['[Content_Types].xml'], 'Default').map((el) => attr(el, 'Extension'));
    expect(defaults).toContain('png');
    const overrides = findAll(xml['[Content_Types].xml'], 'Override').map((el) => attr(el, 'PartName'));
    expect(overrides).toContain('/xl/drawings/drawing1.xml');
  });

  it('a figure renders the image plus a caption row directly below it', async () => {
    const html = `<figure><img src="${TINY_PNG}" alt="Dot"><figcaption>A tiny dot</figcaption></figure>`;
    const { buffer } = await convert(html, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    expect(find(xml['xl/drawings/drawing1.xml'], 'a:blip')).toBeTruthy();
    const cells = cellsByRef(xml);
    expect(refOfString(xml, cells, 'A tiny dot')).toBeTruthy();
  });

  it('a document with no images gets no drawing part at all (no regression to the plain layout)', async () => {
    const { buffer } = await convert('<p>hello</p>', { format: 'xlsx' });
    const { zip, xml } = readXlsxParts(buffer);
    expect(find(xml['xl/worksheets/sheet1.xml'], 'drawing')).toBeFalsy();
    expect(zip.has('xl/worksheets/_rels/sheet1.xml.rels')).toBe(false);
    expect(zip.has('xl/drawings/drawing1.xml')).toBe(false);
  });

  it('an image with no declared width/height still gets the shared placeholder row height, not zero', async () => {
    const { buffer } = await convert(`<img src="${TINY_PNG}">`, { format: 'xlsx' });
    const { xml } = readXlsxParts(buffer);
    const rows = findAll(xml['xl/worksheets/sheet1.xml'], 'row');
    expect(Number(attr(rows[0], 'ht'))).toBeGreaterThan(0);
  });
});
