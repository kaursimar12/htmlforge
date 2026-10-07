import { describe, expect, it } from 'vitest';
import { parseHtml } from '../src/html-parser/index.js';
import { tokenize } from '../src/html-parser/tokenizer.js';

describe('tokenizer', () => {
  it('emits open/close/text tokens', () => {
    const tokens = tokenize('<p>hello</p>');
    expect(tokens).toMatchObject([
      { type: 'open', name: 'p' },
      { type: 'text', value: 'hello' },
      { type: 'close', name: 'p' },
    ]);
  });

  it('handles void elements (self-closing)', () => {
    const tokens = tokenize('<br>');
    expect(tokens[0]).toMatchObject({ type: 'open', name: 'br', selfClosing: true });
  });

  it('parses attributes', () => {
    const tokens = tokenize('<a href="/path" title="t">x</a>');
    expect(tokens[0]).toMatchObject({ type: 'open', name: 'a', attrs: { href: '/path', title: 't' } });
  });

  it('decodes HTML entities in text', () => {
    const tokens = tokenize('<p>a &amp; b</p>');
    expect(tokens[1]).toMatchObject({ type: 'text', value: 'a & b' });
  });

  it('skips script content (raw-text element)', () => {
    const tokens = tokenize('<p>a</p><script>alert(1)</script><p>b</p>');
    const texts = tokens.filter((t) => t.type === 'text').map((t) => t.value);
    expect(texts).toContain('a');
    expect(texts).toContain('b');
    expect(texts).toContain('alert(1)');
    const openNames = tokens.filter((t) => t.type === 'open').map((t) => t.name);
    expect(openNames).not.toContain('alert');
  });
});

describe('parseHtml', () => {
  it('preserves a real space between two adjacent inline elements inside a block container (e.g. <td>)', () => {
    // A very common real-world pattern: `<span class="bold">Bold</span>
    // <span class="italic">Italic</span>` inside a <td> — the whitespace
    // between the spans is significant (renders as a real space in a
    // browser) even though <td> itself is block-level; it must not be
    // dropped just because the *container* happens to be block-level.
    const doc = parseHtml('<table><tr><td><b>Bold</b> <i>Italic</i></td></tr></table>');
    const cell = doc.children[0].rows[0].cells[0];
    expect(cell.children).toMatchObject([
      { type: 'strong', children: [{ type: 'text', value: 'Bold' }] },
      { type: 'text', value: ' ' },
      { type: 'emphasis', children: [{ type: 'text', value: 'Italic' }] },
    ]);
  });

  it('still drops purely insignificant whitespace between two block-level siblings', () => {
    const doc = parseHtml('<div><p>A</p>\n  <p>B</p></div>');
    expect(doc.children).toHaveLength(2);
    expect(doc.children[0]).toMatchObject({ type: 'paragraph' });
    expect(doc.children[1]).toMatchObject({ type: 'paragraph' });
  });

  it('preserves a space between inline elements directly inside a <div>, not just <td>', () => {
    const doc = parseHtml('<div><b>A</b> <i>B</i></div>');
    expect(doc.children[0]).toMatchObject({
      type: 'paragraph',
      children: [
        { type: 'strong', children: [{ type: 'text', value: 'A' }] },
        { type: 'text', value: ' ' },
        { type: 'emphasis', children: [{ type: 'text', value: 'B' }] },
      ],
    });
  });

  it('maps h1 to heading level 1', () => {
    const doc = parseHtml('<h1>Hello</h1>');
    expect(doc.children[0]).toMatchObject({ type: 'heading', level: 1 });
  });

  it('maps p to paragraph', () => {
    const doc = parseHtml('<p>text</p>');
    expect(doc.children[0]).toMatchObject({ type: 'paragraph' });
  });

  it('maps a standalone <button> to a paragraph flagged shrinkToFit, with its own background/padding', () => {
    const doc = parseHtml('<button style="background:#1F9B7A;color:#fff;padding:14px 30px;font-weight:700">Download Word</button>');
    expect(doc.children[0]).toMatchObject({
      type: 'paragraph',
      shrinkToFit: true,
      backgroundColor: '1F9B7A',
      padding: { top: 10.5, right: 22.5, bottom: 10.5, left: 22.5 },
    });
  });

  it("a <button> inside a wrapping <div> keeps its own background/padding; the div's own don't override (shrinkToFit still true)", () => {
    const html = '<div style="text-align:center;padding:28px 0;background:#fff">' +
      '<button style="background:#1F9B7A;padding:14px 30px">Go</button></div>';
    const doc = parseHtml(html);
    expect(doc.children[0]).toMatchObject({
      type: 'paragraph',
      shrinkToFit: true,
      align: 'center', // inherited from the div, since the button itself declared none
      backgroundColor: '1F9B7A', // the button's own, not the div's white
      padding: { top: 10.5, left: 22.5 }, // the button's own, not the div's
    });
  });

  it('maps ul/li to list', () => {
    const doc = parseHtml('<ul><li>A</li><li>B</li></ul>');
    const list = doc.children[0];
    expect(list).toMatchObject({ type: 'list', ordered: false });
    expect(list.items).toHaveLength(2);
  });

  it('maps ol to ordered list', () => {
    const doc = parseHtml('<ol><li>1</li></ol>');
    expect(doc.children[0]).toMatchObject({ type: 'list', ordered: true });
  });

  it('maps table', () => {
    const doc = parseHtml('<table><tr><th>A</th></tr><tr><td>B</td></tr></table>');
    expect(doc.children[0]).toMatchObject({ type: 'table' });
  });

  it('maps strong to strong inline', () => {
    const doc = parseHtml('<p><strong>bold</strong></p>');
    const para = doc.children[0];
    expect(para.children[0]).toMatchObject({ type: 'strong' });
  });

  it('maps a to link', () => {
    const doc = parseHtml('<p><a href="/x">link</a></p>');
    const para = doc.children[0];
    expect(para.children[0]).toMatchObject({ type: 'link', href: '/x' });
  });

  it('drops script tags', () => {
    const doc = parseHtml('<script>alert(1)</script><p>hello</p>');
    expect(doc.children).toHaveLength(1);
    expect(doc.children[0]?.type).toBe('paragraph');
  });

  it('extracts metadata from head', () => {
    const doc = parseHtml('<html lang="fr"><head><title>My Doc</title></head><body><p>hi</p></body></html>');
    expect(doc.metadata?.title).toBe('My Doc');
    expect(doc.metadata?.language).toBe('fr');
  });

  it('maps blockquote to quote', () => {
    const doc = parseHtml('<blockquote><p>quote</p></blockquote>');
    expect(doc.children[0]).toMatchObject({ type: 'quote' });
  });

  it('maps pre>code to code block', () => {
    const doc = parseHtml('<pre><code class="language-js">const x = 1;</code></pre>');
    expect(doc.children[0]).toMatchObject({ type: 'code', language: 'js' });
  });

  it('a <pre style="background:...;color:..."> (a dark-theme code block) carries both onto the code node', () => {
    const doc = parseHtml('<pre style="background:#272822;color:#eee">const x = 1;</pre>');
    expect(doc.children[0]).toMatchObject({ type: 'code', backgroundColor: '272822', color: 'EEEEEE' });
  });

  it('a <pre> with no declared style has no backgroundColor/color fields (exporter default applies)', () => {
    const doc = parseHtml('<pre>const x = 1;</pre>');
    expect(doc.children[0].backgroundColor).toBeUndefined();
    expect(doc.children[0].color).toBeUndefined();
  });

  it('a <div style="white-space:pre"> not using the <pre> tag is still recognized as a code block, preserving its literal line breaks/blank lines', () => {
    const doc = parseHtml('<div style="background:#1e293b;color:#fff;white-space:pre">line one\n\nline two</div>');
    expect(doc.children[0]).toMatchObject({
      type: 'code',
      value: 'line one\n\nline two',
      backgroundColor: '1E293B',
      color: 'FFFFFF',
    });
  });

  it('white-space:pre-wrap/pre-line are recognized the same way, but plain (unset) white-space is not', () => {
    expect(parseHtml('<div style="white-space:pre-wrap">a\nb</div>').children[0].type).toBe('code');
    expect(parseHtml('<div style="white-space:pre-line">a\nb</div>').children[0].type).toBe('code');
    expect(parseHtml('<div>a\nb</div>').children[0].type).toBe('paragraph');
  });

  it('maps section to section node', () => {
    const doc = parseHtml('<section><p>x</p></section>');
    expect(doc.children[0]).toMatchObject({ type: 'section' });
  });

  it('maps hr to thematicBreak', () => {
    const doc = parseHtml('<hr>');
    expect(doc.children[0]).toMatchObject({ type: 'thematicBreak' });
  });

  it('maps hr.page-break to pageBreak', () => {
    const doc = parseHtml('<hr class="page-break">');
    expect(doc.children[0]).toMatchObject({ type: 'pageBreak' });
  });

  it('a table cell containing a nested table maps to block content, not flattened text', () => {
    const doc = parseHtml('<table><tr><td><table><tr><td>Inner</td></tr></table></td></tr></table>');
    const outerCell = doc.children[0].rows[0].cells[0];
    expect(outerCell.children).toHaveLength(1);
    expect(outerCell.children[0]).toMatchObject({ type: 'table' });
    expect(outerCell.children[0].rows[0].cells[0].children[0]).toMatchObject({ type: 'text', value: 'Inner' });
  });

  it('a table cell with only phrasing content still maps inline (no regression)', () => {
    const doc = parseHtml('<table><tr><td><strong>bold</strong> text</td></tr></table>');
    const cell = doc.children[0].rows[0].cells[0];
    expect(cell.children[0]).toMatchObject({ type: 'strong' });
    expect(cell.children[1]).toMatchObject({ type: 'text', value: ' text' });
  });

  it('a table cell containing a paragraph maps to block content', () => {
    const doc = parseHtml('<table><tr><td><p>Para one</p><p>Para two</p></td></tr></table>');
    const cell = doc.children[0].rows[0].cells[0];
    expect(cell.children).toHaveLength(2);
    expect(cell.children[0]).toMatchObject({ type: 'paragraph' });
    expect(cell.children[1]).toMatchObject({ type: 'paragraph' });
  });

  it('a <p style="text-align:center"> carries align on the paragraph node', () => {
    const doc = parseHtml('<p style="text-align:center">hi</p>');
    expect(doc.children[0]).toMatchObject({ type: 'paragraph', align: 'center' });
  });

  it('a heading with text-align carries align', () => {
    const doc = parseHtml('<h2 style="text-align:right">Title</h2>');
    expect(doc.children[0]).toMatchObject({ type: 'heading', level: 2, align: 'right' });
  });

  it('a <p style="color:...">  wraps its content in a styledSpan with the color', () => {
    const doc = parseHtml('<p style="color:#ff0000">red text</p>');
    const span = doc.children[0].children[0];
    expect(span).toMatchObject({ type: 'styledSpan', style: { color: 'FF0000' } });
    expect(span.children[0]).toMatchObject({ type: 'text', value: 'red text' });
  });

  it('a <span style="font-weight:bold"> inline wraps in a styledSpan', () => {
    const doc = parseHtml('<p>plain <span style="font-weight:bold">bold</span></p>');
    const para = doc.children[0];
    expect(para.children[1]).toMatchObject({ type: 'styledSpan', style: { bold: true } });
  });

  it('an element with no recognized style produces no styledSpan wrapper', () => {
    const doc = parseHtml('<p style="margin:10px">plain</p>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'text', value: 'plain' });
  });

  it('an <svg> wrapping an embedded base64 <image> maps to a real ImageNode', () => {
    const html = '<svg width="100" height="50"><image xlink:href="data:image/png;base64,AAAA"/></svg>';
    const doc = parseHtml(html);
    expect(doc.children[0]).toMatchObject({ type: 'image', src: 'data:image/png;base64,AAAA' });
  });

  it('an <svg> with a nested pattern/use/image (design-tool export shape) still resolves the embedded image', () => {
    const html =
      '<svg width="209" height="56" viewBox="0 0 209 56">' +
      '<rect width="208" height="56" fill="url(#p1)"/>' +
      '<defs><pattern id="p1"><use xlink:href="#img1" transform="scale(0.01 0.01)"/></pattern>' +
      '<image id="img1" xlink:href="data:image/png;base64,BBBB"/></defs>' +
      '</svg>';
    const doc = parseHtml(html);
    expect(doc.children[0]).toMatchObject({ type: 'image', src: 'data:image/png;base64,BBBB' });
  });

  it('an <svg> with no embedded raster image (pure vector shape) maps to a real svgShapes node', () => {
    const doc = parseHtml('<svg><rect width="10" height="10"/></svg>');
    expect(doc.children[0]).toMatchObject({ type: 'svgShapes', shapes: [{ kind: 'rect', width: 10, height: 10 }] });
  });

  it('an <svg> with neither an embedded raster nor any recognized vector shape (e.g. only a <path>) still drops silently', () => {
    const doc = parseHtml('<svg><path d="M10 10 L90 90"/></svg>');
    expect(doc.children).toHaveLength(0);
  });

  describe('SVG vector shapes (svgShapes)', () => {
    it('extracts rect/circle/ellipse/line/polygon/polyline with their own geometry', () => {
      const html = `<svg width="200" height="120">
        <rect x="10" y="20" width="30" height="40" fill="#2563eb"/>
        <circle cx="50" cy="60" r="15" fill="orange"/>
        <ellipse cx="70" cy="80" rx="12" ry="8" fill="green"/>
        <line x1="0" y1="0" x2="100" y2="100" stroke="black"/>
        <polygon points="0,0 10,0 5,10" fill="red"/>
        <polyline points="0,0 10,10 20,0" stroke="blue"/>
      </svg>`;
      const doc = parseHtml(html);
      expect(doc.children[0].type).toBe('svgShapes');
      expect(doc.children[0].widthPx).toBe(200);
      expect(doc.children[0].heightPx).toBe(120);
      const [rect, circle, ellipse, line, polygon, polyline] = doc.children[0].shapes;
      expect(rect).toMatchObject({ kind: 'rect', x: 10, y: 20, width: 30, height: 40, fill: '2563EB' });
      expect(circle).toMatchObject({ kind: 'circle', cx: 50, cy: 60, r: 15, fill: 'FFA500' });
      expect(ellipse).toMatchObject({ kind: 'ellipse', cx: 70, cy: 80, rx: 12, ry: 8, fill: '008000' });
      expect(line).toMatchObject({ kind: 'line', x1: 0, y1: 0, x2: 100, y2: 100, stroke: '000000' });
      expect(polygon).toMatchObject({ kind: 'polygon', points: [[0, 0], [10, 0], [5, 10]], fill: 'FF0000' });
      expect(polyline).toMatchObject({ kind: 'polyline', points: [[0, 0], [10, 10], [20, 0]], stroke: '0000FF' });
    });

    it('fill defaults to black when unspecified; an explicit fill="none" means no fill at all', () => {
      const doc = parseHtml('<svg><rect width="10" height="10"/><rect width="10" height="10" fill="none"/></svg>');
      expect(doc.children[0].shapes[0].fill).toBe('000000');
      expect(doc.children[0].shapes[1].fill).toBeUndefined();
    });

    it('stroke defaults to none (undefined), not a visible color', () => {
      const doc = parseHtml('<svg><rect width="10" height="10"/></svg>');
      expect(doc.children[0].shapes[0].stroke).toBeUndefined();
    });

    it('a style attribute wins over a same-named presentation attribute', () => {
      const doc = parseHtml('<svg><rect width="10" height="10" fill="red" style="fill:blue"/></svg>');
      expect(doc.children[0].shapes[0].fill).toBe('0000FF');
    });

    it('stroke-width defaults to 1 when a stroke is set but no width is declared, and is omitted with no stroke', () => {
      const withStroke = parseHtml('<svg><rect width="10" height="10" stroke="red"/></svg>');
      expect(withStroke.children[0].shapes[0].strokeWidthPx).toBe(1);
      const withoutStroke = parseHtml('<svg><rect width="10" height="10"/></svg>');
      expect(withoutStroke.children[0].shapes[0].strokeWidthPx).toBeUndefined();
    });

    it('a viewBox rescales every shape into the declared width/height box', () => {
      // viewBox is a 24x24 icon coordinate space, displayed at 48x48 — every
      // coordinate should double.
      const html = '<svg width="48" height="48" viewBox="0 0 24 24"><rect x="2" y="2" width="10" height="10" fill="red"/><circle cx="12" cy="12" r="5" fill="blue"/></svg>';
      const doc = parseHtml(html);
      const [rect, circle] = doc.children[0].shapes;
      expect(rect).toMatchObject({ x: 4, y: 4, width: 20, height: 20 });
      expect(circle).toMatchObject({ cx: 24, cy: 24, r: 10 });
    });

    it('a viewBox with a non-zero origin offsets every shape accordingly', () => {
      const html = '<svg width="100" height="100" viewBox="10 10 100 100"><rect x="10" y="10" width="20" height="20" fill="red"/></svg>';
      const doc = parseHtml(html);
      expect(doc.children[0].shapes[0]).toMatchObject({ x: 0, y: 0, width: 20, height: 20 });
    });

    it('a <g> with no transform is still walked into; a <g transform="..."> is skipped entirely', () => {
      const grouped = parseHtml('<svg><g><rect width="10" height="10" fill="red"/></g></svg>');
      expect(grouped.children[0].shapes).toHaveLength(1);
      const transformed = parseHtml('<svg><g transform="translate(5,5)"><rect width="10" height="10" fill="red"/></g></svg>');
      expect(transformed.children).toHaveLength(0);
    });

    it('a zero-size rect/circle/ellipse or a <line>/<polygon> with too few points is dropped, not carried as an invisible shape', () => {
      const doc = parseHtml('<svg><rect width="0" height="10" fill="red"/><polygon points="1,1" fill="red"/><rect width="10" height="10" fill="red"/></svg>');
      expect(doc.children[0].shapes).toHaveLength(1);
    });

    it('width/height default to the SVG spec default (300x150) when neither an attribute nor a viewBox declares a size', () => {
      const doc = parseHtml('<svg><rect width="10" height="10" fill="red"/></svg>');
      expect(doc.children[0].widthPx).toBe(300);
      expect(doc.children[0].heightPx).toBe(150);
    });
  });

  it('a <table style="border:1px solid #e5e7eb"> carries a parsed border on the table node', () => {
    const doc = parseHtml('<table style="border:1px solid #e5e7eb"><tr><td>x</td></tr></table>');
    expect(doc.children[0].border).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
  });

  it('a <table> with no border style has no border field', () => {
    const doc = parseHtml('<table><tr><td>x</td></tr></table>');
    expect(doc.children[0].border).toBeUndefined();
  });

  it('a <td style="background-color:...">  carries backgroundColor on the cell node', () => {
    const doc = parseHtml('<table><tr><td style="background-color:#2563eb">&nbsp;</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].backgroundColor).toBe('2563EB');
  });

  it('an <img width height> carries its intrinsic size onto the ImageNode', () => {
    const doc = parseHtml('<img src="a.png" width="209" height="56">');
    expect(doc.children[0]).toMatchObject({ type: 'image', width: 209, height: 56 });
  });

  it('an <svg width height> carries its size onto the extracted ImageNode', () => {
    const html = '<svg width="209" height="56"><image xlink:href="data:image/png;base64,AAAA"/></svg>';
    const doc = parseHtml(html);
    expect(doc.children[0]).toMatchObject({ type: 'image', width: 209, height: 56 });
  });

  it('an <img> with no width/height attributes has undefined dimensions (no guessing)', () => {
    const doc = parseHtml('<img src="a.png">');
    expect(doc.children[0].width).toBeUndefined();
    expect(doc.children[0].height).toBeUndefined();
  });

  it('an <img style="width:...;height:...px"> carries its CSS size onto the ImageNode (no width/height attributes at all)', () => {
    const doc = parseHtml('<img src="a.png" style="width:80px;height:80px;border:1px solid #999">');
    expect(doc.children[0]).toMatchObject({ type: 'image', width: 80, height: 80 });
  });

  it('an <img> with BOTH a width/height attribute and a conflicting CSS size uses the CSS size, matching real browser cascade precedence', () => {
    const doc = parseHtml('<img src="a.png" width="209" height="56" style="width:80px;height:80px">');
    expect(doc.children[0]).toMatchObject({ type: 'image', width: 80, height: 80 });
  });

  it('a <td width="52%"> carries widthPercent on the cell node', () => {
    const doc = parseHtml('<table><tr><td width="52%">A</td><td width="48%">B</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].widthPercent).toBe(52);
    expect(doc.children[0].rows[0].cells[1].widthPercent).toBe(48);
  });

  it('a <td> with no width attribute/style has no widthPercent', () => {
    const doc = parseHtml('<table><tr><td>A</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].widthPercent).toBeUndefined();
  });

  it('a <td style="height:4px"> carries heightPx onto the row', () => {
    const doc = parseHtml('<table><tr><td style="height:4px">&nbsp;</td></tr></table>');
    expect(doc.children[0].rows[0].heightPx).toBe(4);
  });

  it('a row with no explicit height has no heightPx', () => {
    const doc = parseHtml('<table><tr><td>A</td></tr></table>');
    expect(doc.children[0].rows[0].heightPx).toBeUndefined();
  });

  it('a <td style="border-bottom:1px solid #hex"> carries borderBottom on the cell, independent of any table-level border', () => {
    const doc = parseHtml('<table><tr><td style="border-bottom:1px solid #e5e7eb">&nbsp;</td></tr></table>');
    const cell = doc.children[0].rows[0].cells[0];
    expect(cell.borderBottom).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
    expect(cell.borderTop).toBeUndefined();
  });

  it('a <td> with no border style has no cell border fields', () => {
    const doc = parseHtml('<table><tr><td>A</td></tr></table>');
    const cell = doc.children[0].rows[0].cells[0];
    expect(cell.borderTop).toBeUndefined();
    expect(cell.borderBottom).toBeUndefined();
    expect(cell.borderLeft).toBeUndefined();
    expect(cell.borderRight).toBeUndefined();
  });

  it('a <td style="border:none"> carries noBorder:true, distinct from simply having no border declared', () => {
    const doc = parseHtml('<table><tr><td style="border:none">A</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].noBorder).toBe(true);
  });

  it('a <td> with no border declared has no noBorder field (undefined, not false)', () => {
    const doc = parseHtml('<table><tr><td>A</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].noBorder).toBeUndefined();
  });

  it('a <td style="padding:8px"> carries padding on the cell node', () => {
    const doc = parseHtml('<table><tr><td style="padding:8px">A</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].padding).toEqual({ top: 6, right: 6, bottom: 6, left: 6 });
  });

  it('a <td> with no padding style has no padding field', () => {
    const doc = parseHtml('<table><tr><td>A</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].padding).toBeUndefined();
  });

  it('a <th> is center-aligned by default even with no CSS at all, a <td> is not', () => {
    // Mixed th/td in one row isn't auto-promoted to a head row (that only
    // happens when every cell is a <th>), so both land in .rows[0] here.
    const doc = parseHtml('<table><tr><th>Head</th><td>Data</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].align).toBe('center');
    expect(doc.children[0].rows[0].cells[1].align).toBeUndefined();
  });

  it('an explicit text-align on a <th> overrides its center default', () => {
    const doc = parseHtml('<table><tr><th style="text-align:left">Head</th></tr></table>');
    expect(doc.children[0].head[0].cells[0].align).toBe('left');
  });

  it('a <td style="vertical-align:..."> carries verticalAlign on the cell, "middle" mapped to "center"', () => {
    const doc = parseHtml('<table><tr><td style="vertical-align:top">A</td><td style="vertical-align:middle">B</td></tr></table>');
    const cells = doc.children[0].rows[0].cells;
    expect(cells[0].verticalAlign).toBe('top');
    expect(cells[1].verticalAlign).toBe('center');
  });

  it('a <td> with no vertical-align has no verticalAlign field', () => {
    const doc = parseHtml('<table><tr><td>A</td></tr></table>');
    expect(doc.children[0].rows[0].cells[0].verticalAlign).toBeUndefined();
  });

  it('maps <u> to a styledSpan with underline', () => {
    const doc = parseHtml('<p><u>Underline</u></p>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { underline: true } });
  });

  it.each(['s', 'strike', 'del'])('maps <%s> to a styledSpan with strikethrough', (tag) => {
    const doc = parseHtml(`<p><${tag}>Gone</${tag}></p>`);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { strikethrough: true } });
  });

  it('maps <ins> to a styledSpan with underline', () => {
    const doc = parseHtml('<p><ins>Added</ins></p>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { underline: true } });
  });

  it('maps <mark> to a styledSpan with a yellow background', () => {
    const doc = parseHtml('<p><mark>Marked</mark></p>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { backgroundColor: 'FFFF00' } });
  });

  it('maps <sub> to a styledSpan with subscript verticalAlign', () => {
    const doc = parseHtml('<p>H<sub>2</sub>O</p>');
    expect(doc.children[0].children[1]).toMatchObject({ type: 'styledSpan', style: { verticalAlign: 'subscript' } });
  });

  it('maps <sup> to a styledSpan with superscript verticalAlign', () => {
    const doc = parseHtml('<p>x<sup>2</sup></p>');
    expect(doc.children[0].children[1]).toMatchObject({ type: 'styledSpan', style: { verticalAlign: 'superscript' } });
  });

  it('maps <kbd> to a codeSpan, same as <code>', () => {
    const doc = parseHtml('<p><kbd>Ctrl+C</kbd></p>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'codeSpan', value: 'Ctrl+C' });
  });

  it.each(['var', 'cite'])('maps <%s> to a styledSpan with italic', (tag) => {
    const doc = parseHtml(`<p><${tag}>x</${tag}></p>`);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { italic: true } });
  });

  it('maps <q> to its content surrounded by quotation marks, as three separate inline nodes', () => {
    const doc = parseHtml('<p><q>Quote</q></p>');
    const para = doc.children[0];
    expect(para.children.map((n) => n.value)).toEqual(['"', 'Quote', '"']);
  });

  it('a <span style="background:yellow"> (shorthand, not background-color) still resolves as a highlight', () => {
    const doc = parseHtml('<p><span style="background:yellow">Hi</span></p>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { backgroundColor: 'FFFF00' } });
  });

  it('<abbr> keeps its text content but drops the title (no static-document equivalent for a tooltip)', () => {
    const doc = parseHtml('<p><abbr title="HyperText Markup Language">HTML</abbr></p>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'text', value: 'HTML' });
  });

  it('a class from a <style> block resolves the same as an equivalent inline style', () => {
    const html = '<html><head><style>.large{font-size:28px}</style></head><body><p class="large">Big</p></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { fontSizePt: 21 } });
  });

  it('an element with multiple classes merges every matching rule', () => {
    const html =
      '<html><head><style>.bold{font-weight:bold}.red{color:red}</style></head>' +
      '<body><p class="bold red">Hi</p></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { bold: true, color: 'FF0000' } });
  });

  it('inline style overrides a class rule for the same property', () => {
    const html =
      '<html><head><style>.blue{color:blue}</style></head>' +
      '<body><p class="blue" style="color:red">Hi</p></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { color: 'FF0000' } });
  });

  it('a later-defined class rule wins over an earlier one for the same property, regardless of class= attribute order', () => {
    const html =
      '<html><head><style>.a{color:red}.b{color:blue}</style></head>' +
      '<body><p class="b a">Hi</p></body></html>'; // attribute lists "b" first, but "a" is NOT later in the stylesheet
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { color: '0000FF' } });
  });

  it('a bare tag selector in a <style> block resolves the same as an equivalent inline style', () => {
    const html = '<html><head><style>p{color:red}</style></head><body><p>Colored</p></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { color: 'FF0000' } });
  });

  it('a two-part descendant combinator (.class tag) in a <style> block resolves like an inline style on the descendant', () => {
    const html =
      '<html><head><style>.box h1{color:#fff}</style></head>' +
      '<body><div class="box"><h1>Boxed</h1></div></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { color: 'FFFFFF' } });
  });

  it('a descendant combinator only applies within a matching ancestor, not to the same tag elsewhere', () => {
    const html =
      '<html><head><style>.box h1{color:#fff}</style></head>' +
      '<body><div class="box"><h1>Boxed</h1></div><h1>Outside</h1></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { color: 'FFFFFF' } });
    expect(doc.children[1].children[0]).toMatchObject({ type: 'text', value: 'Outside' });
  });

  it('tag+class compounds, 3+-part chains, and pseudo-classes in a <style> block are silently ignored (out of scope), not misapplied', () => {
    const html =
      '<html><head><style>div.box{color:red}.a .b .c{color:green}tr:nth-child(even){color:blue}</style></head>' +
      '<body><div class="box"><p>Plain</p></div></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'text', value: 'Plain' });
  });

  it('a tag selector applies alongside a matching class rule, class declared later in the stylesheet wins ties', () => {
    const html =
      '<html><head><style>p{color:red;font-weight:bold}.note{color:blue}</style></head>' +
      '<body><p class="note">Hi</p></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'styledSpan', style: { bold: true, color: '0000FF' } });
  });

  it('a class with no matching rule in the stylesheet has no effect', () => {
    const html = '<html><head><style>.known{color:red}</style></head><body><p class="unknown">Hi</p></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0].children[0]).toMatchObject({ type: 'text', value: 'Hi' });
  });

  it('a document with no <style> block is unaffected by class attributes', () => {
    const doc = parseHtml('<p class="whatever">Hi</p>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'text', value: 'Hi' });
  });

  it('a <div> with mixed unwrapped inline content and text collapses into one paragraph, not one per child', () => {
    // No wrapping <p> — a very common real-world pattern (CMS/hand-written HTML).
    const doc = parseHtml('<div><strong>Info:</strong> This layout is intentionally Word-friendly.</div>');
    expect(doc.children).toHaveLength(1);
    expect(doc.children[0]).toMatchObject({
      type: 'paragraph',
      children: [
        { type: 'strong', children: [{ type: 'text', value: 'Info:' }] },
        { type: 'text', value: ' This layout is intentionally Word-friendly.' },
      ],
    });
  });

  it('whitespace-only text between two block siblings produces no stray empty paragraph', () => {
    const doc = parseHtml('<div><h2>A</h2>\n  \n<p>B</p></div>');
    expect(doc.children).toHaveLength(2);
    expect(doc.children[0].type).toBe('heading');
    expect(doc.children[1].type).toBe('paragraph');
  });

  it('an inline run adjacent to a real block child only merges with other inline siblings, not across the block', () => {
    const doc = parseHtml('<div>Before <em>text</em><p>Block</p>After <b>text</b></div>');
    expect(doc.children).toHaveLength(3);
    expect(doc.children[0]).toMatchObject({
      type: 'paragraph',
      children: [{ type: 'text', value: 'Before ' }, { type: 'emphasis' }],
    });
    expect(doc.children[1]).toMatchObject({ type: 'paragraph', children: [{ type: 'text', value: 'Block' }] });
    expect(doc.children[2]).toMatchObject({
      type: 'paragraph',
      children: [{ type: 'text', value: 'After ' }, { type: 'strong' }],
    });
  });

  it('a <div style="border:..."> carries border directly on the paragraph it produces', () => {
    const doc = parseHtml('<div style="border:2px solid #444">Boxed</div>');
    expect(doc.children[0]).toMatchObject({
      type: 'paragraph',
      border: { color: '444444', sizeEighthPt: 12, style: 'single' },
    });
  });

  it('a <div style="border-left:..."> (accent stripe) carries only borderLeft, no other side and no uniform border', () => {
    const doc = parseHtml('<div style="border-left:5px solid #1f4e79">Note</div>');
    const p = doc.children[0];
    expect(p.borderLeft).toEqual({ color: '1F4E79', sizeEighthPt: 30, style: 'single' });
    expect(p.borderTop).toBeUndefined();
    expect(p.borderRight).toBeUndefined();
    expect(p.borderBottom).toBeUndefined();
    expect(p.border).toBeUndefined();
  });

  it('a <div style="border:...;border-left:..."> keeps the per-side override independent of the uniform border', () => {
    const doc = parseHtml('<div style="border:1px solid black;border-left:5px solid red">Note</div>');
    const p = doc.children[0];
    expect(p.border).toEqual({ color: '000000', sizeEighthPt: 6, style: 'single' });
    expect(p.borderLeft).toEqual({ color: 'FF0000', sizeEighthPt: 30, style: 'single' });
  });

  it('a <div style="background-color:..."> carries backgroundColor directly on the paragraph', () => {
    const doc = parseHtml('<div style="background-color:#eeeeee">Boxed</div>');
    expect(doc.children[0]).toMatchObject({ type: 'paragraph', backgroundColor: 'EEEEEE' });
  });

  it('a <div style="color:..."> is inherited by a plain child paragraph with no color of its own', () => {
    const doc = parseHtml('<div style="background:#1f4e79;color:#fff"><h1>Title</h1><p>Sub</p></div>');
    expect(doc.children[0]).toMatchObject({
      type: 'heading',
      backgroundColor: '1F4E79',
      children: [{ type: 'styledSpan', style: { color: 'FFFFFF' } }],
    });
    expect(doc.children[1]).toMatchObject({
      type: 'paragraph',
      backgroundColor: '1F4E79',
      children: [{ type: 'styledSpan', style: { color: 'FFFFFF' } }],
    });
  });

  it("a child's own more specific color overrides an ancestor's inherited color", () => {
    // The ancestor's inherited color wraps *outside* the child's own already-baked
    // styledSpan (mapElement runs bottom-up, so the child's own styling exists first)
    // — the inner, more specific color is what a renderer's rProps merge honors last.
    const doc = parseHtml('<div style="color:#fff"><h1 style="color:#000">Title</h1></div>');
    const outer = doc.children[0].children[0];
    expect(outer).toMatchObject({ type: 'styledSpan', style: { color: 'FFFFFF' } });
    expect(outer.children[0]).toMatchObject({ type: 'styledSpan', style: { color: '000000' } });
  });

  it('an outer wrapper\'s own background does not clobber an inner wrapper\'s, only fills gaps; its padding stays on the outer box, not each child', () => {
    const html =
      '<div style="background:#fff;padding:20px">' +
      '<div style="background:#1f4e79;padding:10px"><p>Boxed</p></div>' +
      '<p>Plain</p>' +
      '</div>';
    const doc = parseHtml(html);
    expect(doc.children[0]).toMatchObject({ backgroundColor: '1F4E79', padding: { top: 7.5 } }); // inner wins, its own single-child padding intact
    expect(doc.children[1]).toMatchObject({ backgroundColor: 'FFFFFF' }); // outer's background still highlights the plain sibling
    expect(doc.children[1].padding).toBeUndefined(); // but not its padding — that describes the outer container's own box, not this child's
  });

  it('a <div class="box"> gets its border from a matching <style> class rule', () => {
    const html = '<html><head><style>.box{border:1px dashed red}</style></head><body><div class="box">Hi</div></body></html>';
    const doc = parseHtml(html);
    expect(doc.children[0]).toMatchObject({
      type: 'paragraph',
      border: { color: 'FF0000', style: 'dashed' },
    });
  });

  it('a <div> with an unrepresentable gradient background falls back to the gradient\'s first color stop', () => {
    const doc = parseHtml('<div style="background:linear-gradient(90deg,#ffd,#cdf)">Hi</div>');
    expect(doc.children[0].backgroundColor).toBe('FFFFDD');
  });

  it('a gradient with a "to <direction>" argument still resolves its first color stop', () => {
    const doc = parseHtml('<div style="background:linear-gradient(to right, red, blue)">Hi</div>');
    expect(doc.children[0].backgroundColor).toBe('FF0000');
  });

  it('a radial gradient with a position argument still resolves its first color stop', () => {
    const doc = parseHtml('<div style="background:radial-gradient(circle at center, red, blue)">Hi</div>');
    expect(doc.children[0].backgroundColor).toBe('FF0000');
  });

  it('a gradient whose stops carry a stop-position (e.g. "red 20%") still resolves just the color', () => {
    const doc = parseHtml('<div style="background:linear-gradient(90deg, red 20%, blue 80%)">Hi</div>');
    expect(doc.children[0].backgroundColor).toBe('FF0000');
  });

  it('a gradient with no resolvable color stops has no backgroundColor field', () => {
    const doc = parseHtml('<div style="background:linear-gradient(90deg, url(x.png))">Hi</div>');
    expect(doc.children[0].backgroundColor).toBeUndefined();
  });

  it('a plain <div> with no styling has no border/backgroundColor fields', () => {
    const doc = parseHtml('<div>Hi</div>');
    expect(doc.children[0].border).toBeUndefined();
    expect(doc.children[0].backgroundColor).toBeUndefined();
  });

  it('a <div style="margin:10px 0"> carries both beforePt and afterPt (vertical shorthand)', () => {
    const doc = parseHtml('<div style="margin:10px 0">Hi</div>');
    expect(doc.children[0].margin).toEqual({ beforePt: 7.5, afterPt: 7.5 });
  });

  it('margin-top/margin-bottom longhand override the margin shorthand', () => {
    const doc = parseHtml('<div style="margin:10px 0;margin-bottom:20px">Hi</div>');
    expect(doc.children[0].margin).toEqual({ beforePt: 7.5, afterPt: 15 });
  });

  it('a bare "0" margin value (no unit) resolves to zero, not undefined', () => {
    const doc = parseHtml('<div style="margin-top:0">Hi</div>');
    expect(doc.children[0].margin).toEqual({ beforePt: 0, afterPt: undefined });
  });

  it('a <div> with only margin-bottom set has beforePt undefined', () => {
    const doc = parseHtml('<div style="margin-bottom:12px">Hi</div>');
    expect(doc.children[0].margin).toEqual({ beforePt: undefined, afterPt: 9 });
  });

  it('a <div> with no margin style has no margin field', () => {
    const doc = parseHtml('<div>Hi</div>');
    expect(doc.children[0].margin).toBeUndefined();
  });

  it('a <div style="padding:10px"> carries a uniform padding on all four sides', () => {
    const doc = parseHtml('<div style="padding:10px">Hi</div>');
    expect(doc.children[0].padding).toEqual({ top: 7.5, right: 7.5, bottom: 7.5, left: 7.5 });
  });

  it('a <div style="padding:10px 20px"> expands the 2-value shorthand correctly', () => {
    const doc = parseHtml('<div style="padding:10px 20px">Hi</div>');
    expect(doc.children[0].padding).toEqual({ top: 7.5, right: 15, bottom: 7.5, left: 15 });
  });

  it('padding-top/right/bottom/left longhand overrides the padding shorthand', () => {
    const doc = parseHtml('<div style="padding:10px;padding-left:30px">Hi</div>');
    expect(doc.children[0].padding).toEqual({ top: 7.5, right: 7.5, bottom: 7.5, left: 22.5 });
  });

  it('a <div> with no padding style has no padding field', () => {
    const doc = parseHtml('<div>Hi</div>');
    expect(doc.children[0].padding).toBeUndefined();
  });

  it('<dl><dt>/<dd> map to separate paragraphs, term bolded, not merged into one run', () => {
    const doc = parseHtml('<dl><dt>HTML</dt><dd>Markup language</dd></dl>');
    expect(doc.children).toHaveLength(2);
    expect(doc.children[0]).toMatchObject({
      type: 'paragraph',
      children: [{ type: 'strong', children: [{ type: 'text', value: 'HTML' }] }],
    });
    expect(doc.children[1]).toMatchObject({
      type: 'paragraph',
      children: [{ type: 'text', value: 'Markup language' }],
    });
  });

  it('a <dl> with multiple term/definition pairs keeps each as its own paragraph', () => {
    const doc = parseHtml('<dl><dt>A</dt><dd>1</dd><dt>B</dt><dd>2</dd></dl>');
    expect(doc.children).toHaveLength(4);
  });

  it('an unchecked <input type="checkbox"> maps to a hollow checkbox glyph', () => {
    const doc = parseHtml('<input type="checkbox">');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'text', value: '☐' });
  });

  it('a checked <input type="checkbox" checked> maps to a filled checkbox glyph', () => {
    const doc = parseHtml('<input type="checkbox" checked>');
    expect(doc.children[0].children[0]).toMatchObject({ type: 'text', value: '☑' });
  });

  it('an <input type="text"> (no representable content) produces nothing', () => {
    const doc = parseHtml('<p>Name <input type="text"></p>');
    expect(doc.children[0].children).toEqual([{ type: 'text', value: 'Name ' }]);
  });

  it('a <select> shows its selected <option>\'s text, not every option concatenated', () => {
    const doc = parseHtml('<select><option>One</option><option selected>Two</option></select>');
    expect(doc.children[0].children[0]).toMatchObject({
      type: 'styledSpan', style: { formField: true },
      children: [{ type: 'text', value: 'Two' }],
    });
  });

  it('a <select> with no option marked selected defaults to the first, matching browser behavior', () => {
    const doc = parseHtml('<select><option>One</option><option>Two</option></select>');
    expect(doc.children[0].children[0]).toMatchObject({
      type: 'styledSpan', style: { formField: true },
      children: [{ type: 'text', value: 'One' }],
    });
  });

  it('an empty <select> with no options produces nothing', () => {
    const doc = parseHtml('<select></select>');
    expect(doc.children).toEqual([]);
  });

  it('a <div class="pagebreak"> styled with page-break-before:always inserts a real page break', () => {
    const doc = parseHtml('<div style="page-break-before:always">Content</div>');
    expect(doc.children[0]).toMatchObject({ type: 'pageBreak' });
    expect(doc.children[1]).toMatchObject({ type: 'paragraph' });
  });

  it('break-before:page (the modern CSS property) also inserts a page break', () => {
    const doc = parseHtml('<div style="break-before:page">Content</div>');
    expect(doc.children[0]).toMatchObject({ type: 'pageBreak' });
  });

  it('a <div> with no page-break style produces no pageBreak node', () => {
    const doc = parseHtml('<div>Content</div>');
    expect(doc.children.some((n) => n.type === 'pageBreak')).toBe(false);
  });

  it('an element with an id="..." carries anchorId on the paragraph it produces', () => {
    const doc = parseHtml('<footer id="end">Footer</footer>');
    expect(doc.children[0]).toMatchObject({ type: 'paragraph', anchorId: 'end' });
  });
});

describe('scrollable container detection (preserveScrollableContent)', () => {
  it('a <div> with overflow:auto and a declared height becomes a __scrollableContainer marker instead of unwrapping', () => {
    const doc = parseHtml('<div style="height:50px; overflow:auto"><p>One</p><p>Two</p></div>');
    expect(doc.children).toHaveLength(1);
    expect(doc.children[0]).toMatchObject({
      type: '__scrollableContainer',
      heightPx: 50,
      children: [
        { type: 'paragraph' },
        { type: 'paragraph' },
      ],
    });
  });

  it('overflow:hidden with a declared max-height also qualifies', () => {
    const doc = parseHtml('<div style="max-height:100px; overflow:hidden"><p>Text</p></div>');
    expect(doc.children[0]).toMatchObject({ type: '__scrollableContainer', maxHeightPx: 100 });
  });

  it('overflow-y:scroll (longhand) also qualifies', () => {
    const doc = parseHtml('<div style="height:50px; overflow-y:scroll"><p>Text</p></div>');
    expect(doc.children[0].type).toBe('__scrollableContainer');
  });

  it('a declared height with no overflow style unwraps normally (height alone is not "scrollable")', () => {
    const doc = parseHtml('<div style="height:50px"><p>Text</p></div>');
    expect(doc.children[0]).toMatchObject({ type: 'paragraph' });
  });

  it('overflow:auto with no declared height unwraps normally (overflow alone is not "scrollable")', () => {
    const doc = parseHtml('<div style="overflow:auto"><p>Text</p></div>');
    expect(doc.children[0]).toMatchObject({ type: 'paragraph' });
  });

  it('overflow:visible with a declared height does not qualify', () => {
    const doc = parseHtml('<div style="height:50px; overflow:visible"><p>Text</p></div>');
    expect(doc.children[0]).toMatchObject({ type: 'paragraph' });
  });

  it('a scrollable div nested inside a <section> still gets detected', () => {
    const doc = parseHtml('<section><div style="height:50px; overflow:auto"><p>Text</p></div></section>');
    expect(doc.children[0].children[0].type).toBe('__scrollableContainer');
  });
});
