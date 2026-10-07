import { describe, it, expect } from 'vitest';
import { convert } from '../src/index.js';

async function toJson(html, options) {
  const result = await convert(html, { format: 'json', ...options });
  return JSON.parse(result.buffer.toString('utf8'));
}

describe('JSON exporter', () => {
  it('J1 — mirrors headings, paragraphs, and inline nodes as {type, children}/{type, value}', async () => {
    const data = await toJson('<h1>Title</h1><p>Hello <strong>bold</strong>, <a href="/x">link</a></p>');
    expect(data.content[0]).toEqual({
      type: 'heading',
      level: 1,
      children: [{ type: 'text', value: 'Title' }],
    });
    expect(data.content[1]).toEqual({
      type: 'paragraph',
      children: [
        { type: 'text', value: 'Hello ' },
        { type: 'strong', children: [{ type: 'text', value: 'bold' }] },
        { type: 'text', value: ', ' },
        { type: 'link', href: '/x', children: [{ type: 'text', value: 'link' }] },
      ],
    });
  });

  it('J1 — list items use "children", not the IDM\'s internal "items" key', async () => {
    const data = await toJson('<ul><li>A</li><li>B</li></ul>');
    const list = data.content[0];
    expect(list.type).toBe('list');
    expect(list.ordered).toBe(false);
    expect(list.items).toBeUndefined();
    expect(list.children).toHaveLength(2);
    expect(list.children[0]).toEqual({
      type: 'listItem',
      children: [{ type: 'paragraph', children: [{ type: 'text', value: 'A' }] }],
    });
  });

  it('J1 — ordered list keeps a start offset when present', async () => {
    const data = await toJson('<ol start="3"><li>X</li></ol>');
    expect(data.content[0].ordered).toBe(true);
    expect(data.content[0].start).toBe(3);
  });

  it('J1 — code node keeps language + value, figure normalizes to children', async () => {
    const data = await toJson(
      '<pre><code class="language-js">const x = 1;</code></pre>' +
        '<figure><img src="c.png" alt="Chart"><figcaption>Caption text</figcaption></figure>',
    );
    expect(data.content[0]).toEqual({ type: 'code', language: 'js', value: 'const x = 1;' });
    expect(data.content[1]).toEqual({
      type: 'figure',
      children: [
        { type: 'image', src: 'c.png', alt: 'Chart' },
        { type: 'text', value: 'Caption text' },
      ],
    });
  });

  it('J2 — table becomes flat headers/rows of plain strings', async () => {
    const html =
      '<table><thead><tr><th>Item</th><th>Price</th></tr></thead>' +
      '<tbody><tr><td>Apple</td><td>$1</td></tr><tr><td>Banana</td><td>$2</td></tr></tbody></table>';
    const data = await toJson(html);
    expect(data.content[0]).toEqual({
      type: 'table',
      headers: ['Item', 'Price'],
      rows: [
        ['Apple', '$1'],
        ['Banana', '$2'],
      ],
    });
  });

  it('J2 — table cell formatting is flattened to plain text', async () => {
    const html = '<table><tr><td><strong>bold</strong> and <a href="/x">link</a></td></tr></table>';
    const data = await toJson(html);
    expect(data.content[0].rows[0][0]).toBe('bold and link');
  });

  it('J2 — a <caption> is preserved on the table object', async () => {
    const html = '<table><caption>My Caption</caption><tr><td>x</td></tr></table>';
    const data = await toJson(html);
    expect(data.content[0].caption).toBe('My Caption');
  });

  it('J2 — a table with no thead has an empty headers array, all rows in "rows"', async () => {
    const html = '<table><tr><td>1</td></tr><tr><td>2</td></tr></table>';
    const data = await toJson(html);
    expect(data.content[0].headers).toEqual([]);
    expect(data.content[0].rows).toEqual([['1'], ['2']]);
  });

  it('J3 — metadata from <head> is lifted to the top-level "meta" key', async () => {
    const html = '<html lang="en"><head><title>My Doc</title></head><body><p>x</p></body></html>';
    const data = await toJson(html);
    expect(data.meta.title).toBe('My Doc');
    expect(data.meta.language).toBe('en');
  });

  it('J3 — every image in the document is collected into a top-level "assets" manifest', async () => {
    const html =
      '<p>Intro <img src="inline.png" alt="Inline"></p>' +
      '<img src="https://example.com/photo.jpg" alt="Standalone">' +
      '<figure><img src="chart.svg" alt="Chart"><figcaption>Cap</figcaption></figure>';
    const data = await toJson(html);
    expect(data.assets).toEqual([
      { id: 'asset-1', src: 'inline.png', mimeType: 'image/png' },
      { id: 'asset-2', src: 'https://example.com/photo.jpg', mimeType: 'image/jpeg' },
      { id: 'asset-3', src: 'chart.svg', mimeType: 'image/svg+xml' },
    ]);
  });

  it('J3 — a data: URI asset resolves its mimeType from the URI header, not the extension map', async () => {
    const html = '<img src="data:image/png;base64,AAAA" alt="Inline">';
    const data = await toJson(html);
    expect(data.assets).toEqual([
      { id: 'asset-1', src: 'data:image/png;base64,AAAA', mimeType: 'image/png' },
    ]);
  });

  it('J3 — a document with no images has an empty assets array', async () => {
    const data = await toJson('<p>no images here</p>');
    expect(data.assets).toEqual([]);
  });

  it('an image with alt but a data: URI src shows only the alt — no raw base64 in the content tree', async () => {
    const data = await toJson('<img src="data:image/png;base64,AAAABBBBCCCC" alt="Logo">');
    expect(data.content[0]).toEqual({ type: 'image', alt: 'Logo' });
  });

  it('an image with a data: URI src and no alt drops the src entirely rather than dumping base64', async () => {
    const data = await toJson('<img src="data:image/png;base64,AAAABBBBCCCC">');
    expect(data.content[0]).toEqual({ type: 'image' });
  });

  it('an image with a real (non-data) URL and no alt still shows the src', async () => {
    const data = await toJson('<img src="https://example.com/logo.png">');
    expect(data.content[0]).toEqual({ type: 'image', src: 'https://example.com/logo.png' });
  });

  it('a relative image src is resolved against baseUrl for a readable, portable reference', async () => {
    const data = await toJson('<img src="images/chart.png" alt="Chart">', { baseUrl: 'https://myapp.com/reports/' });
    expect(data.content[0]).toEqual({ type: 'image', src: 'https://myapp.com/reports/images/chart.png', alt: 'Chart' });
  });

  it('the assets manifest (J3) is unaffected — it still carries the real resolvable src', async () => {
    const data = await toJson('<img src="data:image/png;base64,AAAABBBBCCCC" alt="Logo">');
    expect(data.assets[0].src).toBe('data:image/png;base64,AAAABBBBCCCC');
  });

  it('a table cell holding only a data-URI image (no alt) flattens to an empty string, not base64', async () => {
    const html = '<table><tr><td><img src="data:image/png;base64,AAAABBBBCCCC"></td><td>Label</td></tr></table>';
    const data = await toJson(html);
    expect(data.content[0].rows[0]).toEqual(['', 'Label']);
  });

  it('J2 — a cell containing a nested table is kept as real nested structure, not flattened to a string', async () => {
    const html = '<table><tr><td><table><tr><td>Inner</td></tr></table></td></tr></table>';
    const data = await toJson(html);
    expect(data.content[0].rows[0][0]).toEqual({
      type: 'table',
      headers: [],
      rows: [['Inner']],
    });
  });

  it('J2 — a cell with two paragraphs is kept as an array of paragraph nodes', async () => {
    const html = '<table><tr><td><p>One</p><p>Two</p></td></tr></table>';
    const data = await toJson(html);
    expect(data.content[0].rows[0][0]).toEqual([
      { type: 'paragraph', children: [{ type: 'text', value: 'One' }] },
      { type: 'paragraph', children: [{ type: 'text', value: 'Two' }] },
    ]);
  });

  it('acceptance — output parses cleanly and re-stringifying is idempotent', async () => {
    const html = '<h1>T</h1><p>Body</p><table><tr><th>A</th></tr><tr><td>1</td></tr></table>';
    const result = await convert(html, { format: 'json' });
    const text = result.buffer.toString('utf8');

    const parsedOnce = JSON.parse(text);
    const restringified = JSON.stringify(parsedOnce, null, 2) + '\n';
    expect(restringified).toBe(text);
    expect(JSON.parse(restringified)).toEqual(parsedOnce);
  });

  it('returns correct mimeType and extension', async () => {
    const result = await convert('<p>hi</p>', { format: 'json' });
    expect(result.mimeType).toBe('application/json');
    expect(result.extension).toBe('.json');
  });

  it('a styled <span> is preserved as a styledSpan node with its style', async () => {
    const data = await toJson('<p>plain <span style="color:#ff0000;font-weight:bold">red bold</span></p>');
    expect(data.content[0].children[1]).toEqual({
      type: 'styledSpan',
      style: { color: 'FF0000', bold: true, fontWeight: 700 },
      children: [{ type: 'text', value: 'red bold' }],
    });
  });

  it('a paragraph with text-align carries an "align" field', async () => {
    const data = await toJson('<p style="text-align:center">Hi</p>');
    expect(data.content[0].align).toBe('center');
  });

  it('a styled cell value still flattens to a plain string', async () => {
    const html = '<table><tr><td><span style="color:red">Alert</span></td></tr></table>';
    const data = await toJson(html);
    expect(data.content[0].rows[0][0]).toBe('Alert');
  });
});
