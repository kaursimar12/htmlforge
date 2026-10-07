import { describe, it, expect } from 'vitest';
import { convert } from '../src/index.js';

const FIXTURE_HTML = `
<h1>Title</h1>
<h2>Subtitle</h2>
<h3>Section</h3>
<p>Plain paragraph.</p>
<p>Inline: <strong>bold</strong>, <em>italic</em>, <code>inline code</code>, <a href="https://example.com" title="Ex">link</a>.</p>
<ul>
  <li>Apple</li>
  <li>Banana</li>
</ul>
<ol>
  <li>First</li>
  <li>Second</li>
</ol>
<ul>
  <li>Parent
    <ul>
      <li>Child</li>
    </ul>
  </li>
</ul>
<table>
  <thead><tr><th>Item</th><th>Price</th></tr></thead>
  <tbody>
    <tr><td>Apple</td><td>$1</td></tr>
    <tr><td>Banana</td><td>$2</td></tr>
  </tbody>
</table>
<pre><code class="language-js">const x = 1;</code></pre>
<blockquote><p>A wise quote.</p></blockquote>
<img src="https://example.com/photo.png" alt="A photo">
<figure>
  <img src="chart.png" alt="Chart">
  <figcaption>Revenue chart</figcaption>
</figure>
<hr>
`;

const EXPECTED_TXT = `\
TITLE
=====

Subtitle
--------

Section

Plain paragraph.

Inline: bold, italic, \`inline code\`, link (https://example.com).

- Apple
- Banana

1. First
2. Second

- Parent
  - Child

+--------+-------+
| Item   | Price |
+--------+-------+
| Apple  | $1    |
| Banana | $2    |
+--------+-------+

    const x = 1;

> A wise quote.

[image: A photo (https://example.com/photo.png)]

[image: Chart (chart.png)]
Revenue chart

----------------------------------------
`;

describe('Plain Text exporter', () => {
  it('golden snapshot — full fixture matches expected output', async () => {
    const result = await convert(FIXTURE_HTML, { format: 'txt' });
    expect(result.buffer.toString('utf8')).toBe(EXPECTED_TXT);
  });

  it('T1 — H1 is forced to ALL CAPS with = underline', async () => {
    const { buffer } = await convert('<h1>hello world</h1>', { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('HELLO WORLD');
    expect(txt).toContain('===========');
  });

  it('T1 — H2 keeps its authored casing with - underline', async () => {
    const { buffer } = await convert('<h2>hello World</h2>', { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('hello World\n-----------');
  });

  it('T1 — H3+ has no underline', async () => {
    const { buffer } = await convert('<h3>Plain Section</h3><p>next</p>', { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('Plain Section\n\nnext');
  });

  it('T2 — paragraphs are blank-line separated and <br> is a single newline', async () => {
    const { buffer } = await convert('<p>A</p><p>Line1<br>Line2</p>', { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('A\n\nLine1\nLine2');
  });

  it('T3 — unordered and ordered lists, nested indent is 2 spaces', async () => {
    const { buffer } = await convert(
      '<ul><li>X</li><li>Y</li></ul><ol><li>One</li></ol><ul><li>P<ul><li>C</li></ul></li></ul>',
      { format: 'txt' },
    );
    const txt = buffer.toString('utf8');
    expect(txt).toContain('- X');
    expect(txt).toContain('- Y');
    expect(txt).toContain('1. One');
    expect(txt).toContain('- P');
    expect(txt).toContain('  - C');
  });

  it('T4 — table column widths auto-size to the longest cell in that column', async () => {
    const html = '<table><tr><th>A</th><th>Longer</th></tr><tr><td>Wider Value</td><td>x</td></tr></table>';
    const { buffer } = await convert(html, { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('| A           | Longer |');
    expect(txt).toContain('| Wider Value | x      |');
  });

  it('a table used purely for layout (cells hold paragraphs, not short values) renders as flowing text, not an ASCII grid', async () => {
    const html = '<table><tr><td><p>GENERATED FOR</p><p>Generator name</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).not.toContain('+');
    expect(txt).not.toContain('|');
    expect(txt).toBe('GENERATED FOR\n\nGenerator name\n');
  });

  it('a decorative table whose only cell is empty (a spacer/divider row) renders nothing, not an empty box', async () => {
    const html = '<p>Before</p><table><tr><td>&nbsp;</td></tr></table><p>After</p>';
    const { buffer } = await convert(html, { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).not.toContain('+');
    expect(txt).toBe('Before\n\nAfter\n');
  });

  it('a real data table (short inline values per cell) still renders as an ASCII grid, unaffected', async () => {
    const html = '<table><tr><th>Item</th><th>Price</th></tr><tr><td>Apple</td><td>$1</td></tr></table>';
    const { buffer } = await convert(html, { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('| Item  | Price |');
    expect(txt).toContain('| Apple | $1    |');
  });

  it('T5 — link, image, and code span inline rendering', async () => {
    const html = '<p><a href="/x">go</a> <img src="a.png" alt="Alt"> <code>c</code></p>';
    const { buffer } = await convert(html, { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('go (/x)');
    expect(txt).toContain('[image: Alt (a.png)]');
    expect(txt).toContain('`c`');
  });

  it('image with alt but no src (malformed <img>) shows only the alt text', async () => {
    const { buffer } = await convert('<img alt="OnlyAlt">', { format: 'txt' });
    expect(buffer.toString('utf8')).toBe('[image: OnlyAlt]\n');
  });

  it('image with src but no alt attribute shows only the src', async () => {
    const { buffer } = await convert('<img src="only.png">', { format: 'txt' });
    expect(buffer.toString('utf8')).toBe('[image: only.png]\n');
  });

  it('image with neither alt nor src falls back to a bare marker', async () => {
    const { buffer } = await convert('<img>', { format: 'txt' });
    expect(buffer.toString('utf8')).toBe('[image]\n');
  });

  it('figure with an image that has no caption', async () => {
    const { buffer } = await convert('<figure><img src="b.png" alt="B"></figure>', { format: 'txt' });
    expect(buffer.toString('utf8')).toBe('[image: B (b.png)]\n');
  });

  it('T6 — code block is 4-space indented, blockquote prefixes every line with >', async () => {
    const html = '<pre><code>line one\nline two</code></pre><blockquote><p>Quoted.</p></blockquote>';
    const { buffer } = await convert(html, { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('    line one\n    line two');
    expect(txt).toContain('> Quoted.');
  });

  it('a <br> inside a table cell becomes a space, not a broken row', async () => {
    const html = '<table><tr><td>Line1<br>Line2</td></tr></table>';
    const { buffer } = await convert(html, { format: 'txt' });
    const txt = buffer.toString('utf8');
    expect(txt).toContain('| Line1 Line2 |');
  });

  it('a table cell containing a nested table flattens to readable text, not lost content', async () => {
    const html = '<table><tr><td><table><tr><td>Inner</td></tr></table></td></tr></table>';
    const { buffer } = await convert(html, { format: 'txt' });
    expect(buffer.toString('utf8')).toContain('Inner');
  });

  it('returns correct mimeType and extension', async () => {
    const result = await convert('<p>hi</p>', { format: 'txt' });
    expect(result.mimeType).toBe('text/plain');
    expect(result.extension).toBe('.txt');
  });

  it('an image with a data: URI src and an alt shows only "[image: alt]", no base64', async () => {
    const html = '<img src="data:image/png;base64,AAAABBBBCCCC" alt="Logo">';
    const { buffer } = await convert(html, { format: 'txt' });
    expect(buffer.toString('utf8')).toBe('[image: Logo]\n');
  });

  it('an image with a data: URI src and no alt falls back to the bare "[image]" marker, not base64', async () => {
    const { buffer } = await convert('<img src="data:image/png;base64,AAAABBBBCCCC">', { format: 'txt' });
    expect(buffer.toString('utf8')).toBe('[image]\n');
  });

  it('an image with a real (non-data) URL and no alt still shows the src', async () => {
    const { buffer } = await convert('<img src="https://example.com/logo.png">', { format: 'txt' });
    expect(buffer.toString('utf8')).toBe('[image: https://example.com/logo.png]\n');
  });

  it('a relative image src is resolved against baseUrl for a readable, portable reference', async () => {
    const html = '<img src="images/chart.png" alt="Chart">';
    const { buffer } = await convert(html, { format: 'txt', baseUrl: 'https://myapp.com/reports/' });
    expect(buffer.toString('utf8')).toBe('[image: Chart (https://myapp.com/reports/images/chart.png)]\n');
  });

  it('semantic inline tags (u/s/del/ins/mark/sub/sup/kbd/var/cite/q) preserve their text content, no data lost', async () => {
    const html =
      '<p><u>A</u> <s>B</s> <del>C</del> <ins>D</ins> <mark>E</mark> ' +
      'H<sub>2</sub>O x<sup>2</sup> <kbd>Ctrl+C</kbd> <var>x</var> <cite>Book</cite> <q>Quote</q></p>';
    const { buffer } = await convert(html, { format: 'txt' });
    const txt = buffer.toString('utf8');
    for (const text of ['A', 'B', 'C', 'D', 'E', 'H2O', 'x2', 'Ctrl+C', 'x', 'Book', '"Quote"']) {
      expect(txt).toContain(text);
    }
  });

  it('a styled <span> unwraps to its plain text — plain text has no way to show color/bold', async () => {
    const { buffer } = await convert('<p><span style="color:red;font-weight:bold">Alert</span></p>', { format: 'txt' });
    expect(buffer.toString('utf8')).toBe('Alert\n');
  });
});
