import { describe, it, expect } from 'vitest';
import { convert } from '../src/index.js';

describe('CSV exporter', () => {
  it('C1 — thead → header row, tbody → data rows', async () => {
    const html = `
      <table>
        <thead><tr><th>Item</th><th>Price</th><th>Qty</th></tr></thead>
        <tbody>
          <tr><td>Apple</td><td>$1</td><td>3</td></tr>
          <tr><td>Banana</td><td>$2</td><td>5</td></tr>
        </tbody>
      </table>
    `;
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe(
      'Item,Price,Qty\r\nApple,$1,3\r\nBanana,$2,5\r\n',
    );
  });

  it('C1 — no explicit thead but first row is all <th> is treated as header', async () => {
    const html = '<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('A,B\r\n1,2\r\n');
  });

  it('C2 — field containing a comma is quoted', async () => {
    const html = '<table><tr><td>Hello, World</td><td>plain</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('"Hello, World",plain\r\n');
  });

  it('C2 — field containing a double quote is quoted and escaped as ""', async () => {
    const html = '<table><tr><td>She said "hi"</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('"She said ""hi"""\r\n');
  });

  it('C2 — field containing a newline (<br>) is quoted', async () => {
    const html = '<table><tr><td>Line1<br>Line2</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('"Line1\nLine2"\r\n');
  });

  it('C2 — field with no special characters is left unquoted', async () => {
    const html = '<table><tr><td>plain</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('plain\r\n');
  });

  it('C3 — multiple tables are concatenated with a blank line between and a warning is emitted', async () => {
    const html = `
      <table><tr><th>A</th></tr><tr><td>1</td></tr></table>
      <table><tr><th>B</th></tr><tr><td>2</td></tr></table>
    `;
    const { buffer, warnings } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('A\r\n1\r\n\r\nB\r\n2\r\n');
    expect(warnings.some((w) => w.code === 'multiple-tables-concatenated')).toBe(true);
  });

  it('emits a warning and returns empty output when there is no table', async () => {
    const { buffer, warnings } = await convert('<p>no tables here</p>', { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('');
    expect(warnings.some((w) => w.code === 'no-table')).toBe(true);
  });

  it('inline formatting inside cells is flattened to plain text', async () => {
    const html = '<table><tr><td><strong>bold</strong> and <em>italic</em></td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('bold and italic\r\n');
  });

  it('returns correct mimeType and extension', async () => {
    const html = '<table><tr><td>x</td></tr></table>';
    const result = await convert(html, { format: 'csv' });
    expect(result.mimeType).toBe('text/csv');
    expect(result.extension).toBe('.csv');
  });

  it('a heading before a table becomes an unpadded label row', async () => {
    const html = '<h1>Employees</h1><table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('Employees\r\nA,B\r\n1,2\r\n');
  });

  it('a <caption> takes priority over a preceding heading', async () => {
    const html = '<h1>Wrong Label</h1><table><caption>Right Label</caption><tr><td>x</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('Right Label\r\nx\r\n');
  });

  it('each table in a multi-table document gets its own heading as a label (no leaking to the next table)', async () => {
    const html = `
      <h1>Employees</h1>
      <table><tr><th>ID</th></tr><tr><td>E1</td></tr></table>
      <h1>Products</h1>
      <table><tr><th>SKU</th></tr><tr><td>P1</td></tr></table>
    `;
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe(
      'Employees\r\nID\r\nE1\r\n\r\nProducts\r\nSKU\r\nP1\r\n',
    );
  });

  it('a table with no preceding heading and no caption has no label row', async () => {
    const html = '<p>intro</p><table><tr><td>x</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('x\r\n');
  });

  it('a cell containing a nested table flattens to readable text instead of being lost', async () => {
    const html = '<table><tr><td><table><tr><td>Inner</td></tr></table></td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toContain('Inner');
  });

  it('a cell with two paragraphs joins them on separate lines within the quoted field', async () => {
    const html = '<table><tr><td><p>One</p><p>Two</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('"One\nTwo"\r\n');
  });

  it('a cell holding a data-URI image with an alt shows only the alt, no base64', async () => {
    const html = '<table><tr><td><img src="data:image/png;base64,AAAABBBBCCCC" alt="Logo"></td><td>x</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('Logo,x\r\n');
  });

  it('a cell holding a data-URI image with no alt shows an empty field, not base64', async () => {
    const html = '<table><tr><td><img src="data:image/png;base64,AAAABBBBCCCC"></td><td>x</td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe(',x\r\n');
  });

  it('a cell holding an image with a real (non-data) URL and no alt shows the src', async () => {
    const html = '<table><tr><td><img src="https://example.com/logo.png"></td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('https://example.com/logo.png\r\n');
  });

  it('semantic inline tags (u/s/del/ins/mark/sub/sup/kbd/var/cite/q) preserve their text content, no data lost', async () => {
    const html =
      '<table><tr><td><u>A</u> <s>B</s> <del>C</del> <ins>D</ins> <mark>E</mark> ' +
      'H<sub>2</sub>O x<sup>2</sup> <kbd>Ctrl+C</kbd> <var>x</var> <cite>Book</cite> <q>Quote</q></td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    const csv = buffer.toString('utf8');
    for (const text of ['A', 'B', 'C', 'D', 'E', 'H2O', 'x2', 'Ctrl+C', 'x', 'Book', 'Quote']) {
      expect(csv).toContain(text);
    }
  });

  it('a styled <span> in a cell flattens to its plain text, formatting dropped', async () => {
    const html = '<table><tr><td><span style="color:red;font-weight:bold">Alert</span></td></tr></table>';
    const { buffer } = await convert(html, { format: 'csv' });
    expect(buffer.toString('utf8')).toBe('Alert\r\n');
  });
});
