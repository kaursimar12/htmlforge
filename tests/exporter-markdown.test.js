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

const EXPECTED_MD = `\
# Title

## Subtitle

### Section

Plain paragraph.

Inline: **bold**, *italic*, \`inline code\`, [link](https://example.com "Ex").

- Apple
- Banana

1. First
2. Second

- Parent
  - Child

| Item | Price |
| --- | --- |
| Apple | $1 |
| Banana | $2 |

\`\`\`js
const x = 1;
\`\`\`

> A wise quote.

![A photo](https://example.com/photo.png)

![Chart](chart.png)
*Revenue chart*

---
`;

describe('Markdown exporter', () => {
  it('golden snapshot — full fixture matches expected output', async () => {
    const result = await convert(FIXTURE_HTML, { format: 'markdown' });
    expect(result.buffer.toString('utf8')).toBe(EXPECTED_MD);
  });

  it('M1 — heading levels', async () => {
    const { buffer } = await convert('<h1>A</h1><h3>B</h3><h6>C</h6>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('# A');
    expect(md).toContain('### B');
    expect(md).toContain('###### C');
  });

  it('M2 — inline formatting', async () => {
    const { buffer } = await convert('<p><strong>b</strong> <em>i</em> <code>c</code> <a href="/x">l</a></p>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('**b**');
    expect(md).toContain('*i*');
    expect(md).toContain('`c`');
    expect(md).toContain('[l](/x)');
  });

  it('M3 — unordered list', async () => {
    const { buffer } = await convert('<ul><li>X</li><li>Y</li></ul>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('- X');
    expect(md).toContain('- Y');
  });

  it('M3 — ordered list', async () => {
    const { buffer } = await convert('<ol><li>One</li><li>Two</li></ol>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('1. One');
    expect(md).toContain('2. Two');
  });

  it('M3 — nested list indents 2 spaces', async () => {
    const { buffer } = await convert('<ul><li>P<ul><li>C</li></ul></li></ul>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('- P');
    expect(md).toContain('  - C');
  });

  it('M4 — GFM table has separator row', async () => {
    const { buffer } = await convert('<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('| A | B |');
    expect(md).toContain('| --- | --- |');
    expect(md).toContain('| 1 | 2 |');
  });

  it('M5 — fenced code block with language', async () => {
    const { buffer } = await convert('<pre><code class="language-ts">let x = 1;</code></pre>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('```ts\nlet x = 1;\n```');
  });

  it('M5 — blockquote prefixes every line with >', async () => {
    const { buffer } = await convert('<blockquote><p>Quote line.</p></blockquote>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('> Quote line.');
  });

  it('M6 — standalone image', async () => {
    const { buffer } = await convert('<img src="a.png" alt="Alt">', { format: 'markdown' });
    expect(buffer.toString('utf8')).toContain('![Alt](a.png)');
  });

  it('M6 — figure with caption', async () => {
    const { buffer } = await convert('<figure><img src="b.png" alt="B"><figcaption>Cap</figcaption></figure>', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('![B](b.png)');
    expect(md).toContain('*Cap*');
  });

  it('M7 — thematic break → ---', async () => {
    const { buffer } = await convert('<p>A</p><hr><p>B</p>', { format: 'markdown' });
    expect(buffer.toString('utf8')).toContain('---');
  });

  it('a table cell containing a nested table flattens to readable text, not lost content', async () => {
    const html = '<table><tr><td><table><tr><td>Inner</td></tr></table></td></tr></table>';
    const { buffer } = await convert(html, { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('Inner');
  });

  it('a table cell with two paragraphs joins them with <br>', async () => {
    const html = '<table><tr><td><p>One</p><p>Two</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'markdown' });
    expect(buffer.toString('utf8')).toContain('One<br>Two');
  });

  it('returns correct mimeType and extension', async () => {
    const result = await convert('<p>hi</p>', { format: 'markdown' });
    expect(result.mimeType).toBe('text/markdown');
    expect(result.extension).toBe('.md');
  });

  it('an image with a data: URI src and an alt shows only the alt text, no base64', async () => {
    const html = '<img src="data:image/png;base64,AAAABBBBCCCC" alt="Logo">';
    const { buffer } = await convert(html, { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).toContain('Logo');
    expect(md).not.toContain('AAAABBBBCCCC');
    expect(md).not.toContain('![');
  });

  it('an image with a data: URI src and no alt falls back to a generic marker, not base64', async () => {
    const { buffer } = await convert('<img src="data:image/png;base64,AAAABBBBCCCC">', { format: 'markdown' });
    const md = buffer.toString('utf8');
    expect(md).not.toContain('AAAABBBBCCCC');
    expect(md).toContain('(image)');
  });

  it('an image with a real (non-data) URL still uses standard markdown image syntax', async () => {
    const { buffer } = await convert('<img src="https://example.com/logo.png" alt="Logo">', { format: 'markdown' });
    expect(buffer.toString('utf8')).toContain('![Logo](https://example.com/logo.png)');
  });

  it('a relative image src is resolved against baseUrl for a readable, portable reference', async () => {
    const html = '<img src="images/chart.png" alt="Chart">';
    const { buffer } = await convert(html, { format: 'markdown', baseUrl: 'https://myapp.com/reports/' });
    expect(buffer.toString('utf8')).toContain('![Chart](https://myapp.com/reports/images/chart.png)');
  });

  it('with no baseUrl, a relative image src is printed unchanged (no regression)', async () => {
    const html = '<img src="images/chart.png" alt="Chart">';
    const { buffer } = await convert(html, { format: 'markdown' });
    expect(buffer.toString('utf8')).toContain('![Chart](images/chart.png)');
  });

  it('semantic inline tags (u/s/del/ins/mark/sub/sup/kbd/var/cite/q) preserve their text content, no data lost', async () => {
    const html =
      '<p><u>A</u> <s>B</s> <del>C</del> <ins>D</ins> <mark>E</mark> ' +
      'H<sub>2</sub>O x<sup>2</sup> <kbd>Ctrl+C</kbd> <var>x</var> <cite>Book</cite> <q>Quote</q></p>';
    const { buffer } = await convert(html, { format: 'markdown' });
    const md = buffer.toString('utf8');
    for (const text of ['A', 'B', 'C', 'D', 'E', 'H2O', 'x2', 'Ctrl+C', 'x', 'Book', '"Quote"']) {
      expect(md).toContain(text);
    }
  });

  it('a <span style="font-weight:bold"> renders with ** like <strong> (color/background dropped, no equivalent)', async () => {
    const { buffer } = await convert('<p><span style="font-weight:bold;color:red">Alert</span></p>', { format: 'markdown' });
    expect(buffer.toString('utf8')).toContain('**Alert**');
  });

  it('a styled span with no bold/italic degrades to its plain text content, not dropped', async () => {
    const { buffer } = await convert('<p><span style="color:blue">Hello</span></p>', { format: 'markdown' });
    expect(buffer.toString('utf8')).toContain('Hello');
  });
});
