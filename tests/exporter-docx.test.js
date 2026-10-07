import { describe, it, expect } from 'vitest';
import { deflateSync } from 'node:zlib';
import { convert } from '../src/index.js';
import { readDocxParts } from './helpers/docx.js';
import { findAll, find, attr, textContent } from '../src/utils/xml.js';

// A real (not just claimed) PNG of the given size — needed to test the
// "scale down to fit the page" path now that intrinsic pixel size always
// wins over a declared width/height attribute (see image-size.js), so a
// fixture's OWN real dimensions, not an <img width/height> attribute, are
// what must actually be oversized here. Mirrors exporter-pdf.test.js's own
// buildTestPng.
function buildTestPng(width, height) {
  const crc32Table = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();
  const crc32 = (buf) => {
    let c = 0xffffffff;
    for (const byte of buf) c = crc32Table[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const typeBuf = Buffer.from(type, 'ascii');
    const lenBuf = Buffer.alloc(4);
    lenBuf.writeUInt32BE(data.length, 0);
    const crcBuf = Buffer.alloc(4);
    crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
    return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
  };
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // color type: RGB
  const rowBytes = 1 + width * 3;
  const raw = Buffer.alloc(rowBytes * height, 128);
  for (let y = 0; y < height; y++) raw[y * rowBytes] = 0; // filter type 0 (None) per row
  const idat = deflateSync(raw);
  return Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', idat), chunk('IEND', Buffer.alloc(0))]);
}

const REQUIRED_PARTS = [
  '[Content_Types].xml',
  '_rels/.rels',
  'docProps/core.xml',
  'docProps/app.xml',
  'word/document.xml',
  'word/styles.xml',
  'word/numbering.xml',
  'word/_rels/document.xml.rels',
];

describe('DOCX exporter', () => {
  it('returns correct mimeType and extension', async () => {
    const result = await convert('<p>hi</p>', { format: 'docx' });
    expect(result.mimeType).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    );
    expect(result.extension).toBe('.docx');
  });

  it('W1 — produces a valid ZIP with every required OOXML part, each parsing as well-formed XML', async () => {
    const { buffer } = await convert('<p>hi</p>', { format: 'docx' });
    expect(buffer[0]).toBe(0x50); // P
    expect(buffer[1]).toBe(0x4b); // K

    const { zip, xml } = readDocxParts(buffer);
    for (const part of REQUIRED_PARTS) {
      expect(zip.has(part), `missing part: ${part}`).toBe(true);
      expect(xml[part], `did not parse as XML: ${part}`).toBeTruthy();
    }
  });

  it('W1 — [Content_Types].xml declares every part with an Override', async () => {
    const { buffer } = await convert('<p>hi</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const overrides = findAll(xml['[Content_Types].xml'], 'Override').map((el) => attr(el, 'PartName'));
    expect(overrides).toEqual(
      expect.arrayContaining([
        '/word/document.xml',
        '/word/styles.xml',
        '/word/numbering.xml',
        '/docProps/core.xml',
        '/docProps/app.xml',
      ]),
    );
  });

  it('W1 — _rels/.rels points at word/document.xml', async () => {
    const { buffer } = await convert('<p>hi</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const targets = findAll(xml['_rels/.rels'], 'Relationship').map((el) => attr(el, 'Target'));
    expect(targets).toContain('word/document.xml');
  });

  it('W1 — word/document.xml root element declares the w/r/wp/a/pic namespaces', async () => {
    const { buffer } = await convert('<p>hi</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const root = xml['word/document.xml'];
    expect(root.name).toBe('w:document');
    expect(attr(root, 'xmlns:w')).toBe('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    expect(attr(root, 'xmlns:wp')).toBeTruthy();
    expect(attr(root, 'xmlns:a')).toBeTruthy();
    expect(attr(root, 'xmlns:pic')).toBeTruthy();
  });

  it('W1 — word/_rels/document.xml.rels references styles.xml and numbering.xml', async () => {
    const { buffer } = await convert('<p>hi</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const targets = findAll(xml['word/_rels/document.xml.rels'], 'Relationship').map((el) => attr(el, 'Target'));
    expect(targets).toContain('styles.xml');
    expect(targets).toContain('numbering.xml');
  });

  it('W2 — a paragraph becomes <w:p><w:r><w:t>', async () => {
    const { buffer } = await convert('<p>Hello World</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(paragraphs).toHaveLength(1);
    expect(textContent(paragraphs[0])).toBe('Hello World');
    const run = findAll(paragraphs[0], 'w:t')[0];
    expect(attr(run, 'xml:space')).toBe('preserve');
  });

  it('W2 — multiple paragraphs appear in document order', async () => {
    const { buffer } = await convert('<p>First</p><p>Second</p><p>Third</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(paragraphs.map(textContent)).toEqual(['First', 'Second', 'Third']);
  });

  it('W2 — inline formatting not yet supported still renders as plain text (no content loss)', async () => {
    const { buffer } = await convert('<p>Hello <strong>bold</strong> world</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(textContent(paragraphs[0])).toBe('Hello bold world');
  });

  it('W3 — headings 1-6 get w:pStyle Heading1-Heading6, and styles.xml defines all six', async () => {
    const html = '<h1>A</h1><h2>B</h2><h3>C</h3><h4>D</h4><h5>E</h5><h6>F</h6>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    for (let level = 1; level <= 6; level++) {
      const pStyle = find(paragraphs[level - 1], 'w:pStyle');
      expect(attr(pStyle, 'w:val')).toBe(`Heading${level}`);
    }

    const styleIds = findAll(xml['word/styles.xml'], 'w:style').map((el) => attr(el, 'w:styleId'));
    for (let level = 1; level <= 6; level++) {
      expect(styleIds).toContain(`Heading${level}`);
    }
  });

  it('W3 — <hr> becomes a paragraph with a bottom border (no dedicated hr construct in OOXML)', async () => {
    const { buffer } = await convert('<p>A</p><hr><p>B</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(paragraphs).toHaveLength(3);
    expect(find(paragraphs[1], 'w:bottom')).toBeTruthy();
    expect(textContent(paragraphs[1])).toBe('');
  });

  it('W3 — <hr class="page-break"> becomes a run with w:br type="page"', async () => {
    const { buffer } = await convert('<p>A</p><hr class="page-break"><p>B</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const breakEl = find(xml['word/document.xml'], 'w:br');
    expect(attr(breakEl, 'w:type')).toBe('page');
  });

  it('W4 — strong/emphasis become w:b/w:i, nested formatting accumulates on the leaf run', async () => {
    const { buffer } = await convert('<p><strong>bold</strong> <em>italic</em> <strong><em>both</em></strong></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const runs = findAll(xml['word/document.xml'], 'w:r').filter((r) => textContent(r).trim() !== '');
    const boldRun = runs.find((r) => textContent(r) === 'bold');
    const italicRun = runs.find((r) => textContent(r) === 'italic');
    const bothRun = runs.find((r) => textContent(r) === 'both');
    expect(find(boldRun, 'w:b')).toBeTruthy();
    expect(find(italicRun, 'w:i')).toBeTruthy();
    expect(find(bothRun, 'w:b')).toBeTruthy();
    expect(find(bothRun, 'w:i')).toBeTruthy();
  });

  it('W4 — codeSpan gets Courier New + shading', async () => {
    const { buffer } = await convert('<p>run <code>x=1</code></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const runs = findAll(xml['word/document.xml'], 'w:r');
    const codeRun = runs.find((r) => textContent(r) === 'x=1');
    expect(attr(find(codeRun, 'w:rFonts'), 'w:ascii')).toBe('Courier New');
    expect(find(codeRun, 'w:shd')).toBeTruthy();
  });

  it('W4 — blockquote applies the Quote style directly to the real paragraph, not a separate empty stub', async () => {
    const { buffer } = await convert('<blockquote><p>Quoted text.</p></blockquote>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(paragraphs).toHaveLength(1); // regression guard: no spurious empty paragraph
    expect(textContent(paragraphs[0])).toBe('Quoted text.');
    expect(attr(find(paragraphs[0], 'w:pStyle'), 'w:val')).toBe('Quote');
  });

  it('a standalone <button> renders as an autofit, borderless, centered single-cell table (not a full-width shaded paragraph)', async () => {
    const html = '<div style="text-align:center"><button style="background:#1F9B7A;color:#fff;padding:14px 30px">Download Word</button></div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    expect(find(doc, 'w:tbl')).toBeTruthy();
    const tblPr = find(doc, 'w:tblPr');
    expect(attr(find(tblPr, 'w:tblW'), 'w:type')).toBe('auto');
    expect(attr(find(tblPr, 'w:tblLayout'), 'w:type')).toBe('autofit');
    expect(attr(find(tblPr, 'w:jc'), 'w:val')).toBe('center');
    const borders = find(tblPr, 'w:tblBorders');
    expect(attr(find(borders, 'w:top'), 'w:val')).toBe('none');
    const tc = find(doc, 'w:tc');
    expect(attr(find(tc, 'w:shd'), 'w:fill')).toBe('1F9B7A');
    const tcMar = find(tc, 'w:tcMar');
    expect(Number(attr(find(tcMar, 'w:left'), 'w:w'))).toBe(450); // 30px = 22.5pt * 20 twips/pt
    expect(textContent(tc)).toBe('Download Word');
  });

  it("a <button>'s own paragraph zeroes out Normal style's default spacing, so its box height comes only from its own padding", async () => {
    const html = '<button style="background:#1F9B7A;padding:14px 30px">Download Word</button>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tc = find(xml['word/document.xml'], 'w:tc');
    const spacing = find(tc, 'w:spacing');
    expect(Number(attr(spacing, 'w:before'))).toBe(0);
    expect(Number(attr(spacing, 'w:after'))).toBe(0);
  });

  it('a <button> with an explicit border renders it on the table, not "none"', async () => {
    const html = '<button style="border:2px solid #444">Click</button>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const borders = find(xml['word/document.xml'], 'w:tblBorders');
    expect(attr(find(borders, 'w:top'), 'w:val')).toBe('single');
    expect(attr(find(borders, 'w:top'), 'w:color')).toBe('444444');
  });

  it('W4 — a multi-line code block joins lines with w:br inside one shaded, monospace paragraph', async () => {
    const html = '<pre><code>line one\nline two</code></pre>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    const codePara = paragraphs.find((p) => textContent(p).includes('line one'));
    expect(textContent(codePara)).toBe('line oneline two');
    expect(find(codePara, 'w:br')).toBeTruthy();
    expect(find(codePara, 'w:shd')).toBeTruthy();
    expect(attr(find(codePara, 'w:rFonts'), 'w:ascii')).toBe('Courier New');
  });

  it('a <pre style="background:...;color:..."> (dark-theme code block) uses its own colors, not the default light-gray shading', async () => {
    const html = '<pre style="background:#272822;color:#eee">const x = 1;</pre>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(attr(find(p, 'w:shd'), 'w:fill')).toBe('272822');
    expect(attr(find(p, 'w:color'), 'w:val')).toBe('EEEEEE');
  });

  it('a <pre> with no declared colors still gets the default light-gray shading (no regression)', async () => {
    const { buffer } = await convert('<pre>const x = 1;</pre>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(attr(find(p, 'w:shd'), 'w:fill')).toBe('F1F1F1');
    expect(find(p, 'w:color')).toBeFalsy();
  });

  it('inline CSS — color/background/underline/strikethrough/font-size on a styled span map to real run properties', async () => {
    const html = '<p><span style="color:#ff0000;background-color:#ffff00;text-decoration:underline line-through;font-size:18px">styled</span></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'styled');
    expect(attr(find(run, 'w:color'), 'w:val')).toBe('FF0000');
    expect(attr(find(run, 'w:shd'), 'w:fill')).toBe('FFFF00');
    expect(attr(find(run, 'w:u'), 'w:val')).toBe('single');
    expect(find(run, 'w:strike')).toBeTruthy();
    expect(Number(attr(find(run, 'w:sz'), 'w:val'))).toBe(27); // 18px * 0.75 = 13.5pt * 2 half-points, rounded
  });

  it('inline CSS — font-weight/font-style on a <span> (not just <strong>/<em>) produce w:b/w:i', async () => {
    const html = '<p><span style="font-weight:bold;font-style:italic">styled</span></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'styled');
    expect(find(run, 'w:b')).toBeTruthy();
    expect(find(run, 'w:i')).toBeTruthy();
  });

  it('inline CSS — a <p style="text-align:center"> emits w:jc on the paragraph', async () => {
    const { buffer } = await convert('<p style="text-align:center">Centered</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(attr(find(p, 'w:jc'), 'w:val')).toBe('center');
  });

  it('inline CSS — a heading with text-align keeps its Heading style and adds w:jc', async () => {
    const { buffer } = await convert('<h1 style="text-align:right">Title</h1>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(attr(find(p, 'w:pStyle'), 'w:val')).toBe('Heading1');
    expect(attr(find(p, 'w:jc'), 'w:val')).toBe('right');
  });

  it('W5 — an unordered list gets ListBullet style, ilvl 0, and a numId pointing at the bullet abstractNum', async () => {
    const { buffer } = await convert('<ul><li>A</li><li>B</li></ul>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(paragraphs.map(textContent)).toEqual(['A', 'B']);
    for (const p of paragraphs) {
      expect(attr(find(p, 'w:pStyle'), 'w:val')).toBe('ListBullet');
      expect(Number(attr(find(p, 'w:ilvl'), 'w:val'))).toBe(0);
    }
    const numIdA = attr(find(paragraphs[0], 'w:numId'), 'w:val');
    const numIdB = attr(find(paragraphs[1], 'w:numId'), 'w:val');
    expect(numIdB).toBe(numIdA); // same list → same numId for both items

    const numEl = findAll(xml['word/numbering.xml'], 'w:num').find((n) => attr(n, 'w:numId') === numIdA);
    expect(Number(attr(find(numEl, 'w:abstractNumId'), 'w:val'))).toBe(0);
  });

  it('W5 — an ordered list gets ListNumber style and points at the decimal abstractNum', async () => {
    const { buffer } = await convert('<ol><li>One</li></ol>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(attr(find(paragraphs[0], 'w:pStyle'), 'w:val')).toBe('ListNumber');
    const numId = attr(find(paragraphs[0], 'w:numId'), 'w:val');
    const numEl = findAll(xml['word/numbering.xml'], 'w:num').find((n) => attr(n, 'w:numId') === numId);
    expect(Number(attr(find(numEl, 'w:abstractNumId'), 'w:val'))).toBe(1);
  });

  it('W5 — a nested list increases ilvl and registers its own, distinct numId', async () => {
    const html = '<ul><li>Parent<ul><li>Child</li></ul></li></ul>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(paragraphs.map(textContent)).toEqual(['Parent', 'Child']);
    expect(Number(attr(find(paragraphs[0], 'w:ilvl'), 'w:val'))).toBe(0);
    expect(Number(attr(find(paragraphs[1], 'w:ilvl'), 'w:val'))).toBe(1);
    const parentNumId = attr(find(paragraphs[0], 'w:numId'), 'w:val');
    const childNumId = attr(find(paragraphs[1], 'w:numId'), 'w:val');
    expect(childNumId).not.toBe(parentNumId);
  });

  it('W5 — two separate top-level lists are independently numbered (distinct numIds)', async () => {
    const html = '<ul><li>A</li></ul><ul><li>B</li></ul>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    const numIdA = attr(find(paragraphs[0], 'w:numId'), 'w:val');
    const numIdB = attr(find(paragraphs[1], 'w:numId'), 'w:val');
    expect(numIdA).not.toBe(numIdB);
  });

  it('W5 — <ol start="N"> emits a startOverride on the num definition', async () => {
    const { buffer } = await convert('<ol start="3"><li>X</li></ol>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    const numId = attr(find(paragraphs[0], 'w:numId'), 'w:val');
    const numEl = findAll(xml['word/numbering.xml'], 'w:num').find((n) => attr(n, 'w:numId') === numId);
    expect(Number(attr(find(numEl, 'w:startOverride'), 'w:val'))).toBe(3);
  });

  // Every row's total gridSpan must equal the declared column count — this
  // exact invariant is what the prior DOCX attempt violated (mismatched cell
  // counts across rows), producing Word's "unreadable content, repair?" prompt.
  function assertColumnCountConsistent(tbl) {
    const columnCount = findAll(tbl, 'w:gridCol').length;
    for (const row of findAll(tbl, 'w:tr')) {
      const total = findAll(row, 'w:tc').reduce((sum, tc) => {
        const span = find(tc, 'w:gridSpan');
        return sum + (span ? Number(attr(span, 'w:val')) : 1);
      }, 0);
      expect(total).toBe(columnCount);
    }
  }

  it('W6 — a simple table gets w:tblGrid, consistent column counts, and bold header cells', async () => {
    const html = '<table><tr><th>Item</th><th>Price</th></tr><tr><td>Apple</td><td>$1</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    expect(findAll(tbl, 'w:gridCol')).toHaveLength(2);
    assertColumnCountConsistent(tbl);

    const rows = findAll(tbl, 'w:tr');
    const headerCells = findAll(rows[0], 'w:tc');
    for (const tc of headerCells) expect(find(tc, 'w:b')).toBeTruthy();
    const bodyCells = findAll(rows[1], 'w:tc');
    for (const tc of bodyCells) expect(find(tc, 'w:b')).toBeFalsy();
    expect(bodyCells.map(textContent)).toEqual(['Apple', '$1']);
  });

  it('W6 — a real data table (plain inline cell values) gets the bordered TableGrid style', async () => {
    const html = '<table><tr><th>Item</th><th>Price</th></tr><tr><td>Apple</td><td>$1</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    expect(attr(find(tbl, 'w:tblStyle'), 'w:val')).toBe('TableGrid');
    expect(find(tbl, 'w:tblBorders')).toBeFalsy();
  });

  it('W6 — a layout table (cells hold block content) renders borderless instead of a boxy TableGrid', async () => {
    const html = '<table><tr><td><p>Label</p><p>Value</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    expect(find(tbl, 'w:tblStyle')).toBeFalsy();
    const borders = find(tbl, 'w:tblBorders');
    expect(borders).toBeTruthy();
    for (const edge of findAll(borders, 'w:top').concat(findAll(borders, 'w:insideH'))) {
      expect(attr(edge, 'w:val')).toBe('none');
    }
  });

  it('W6 — colspan emits w:gridSpan and expands the declared column count', async () => {
    const html = '<table><tr><td colspan="2">AB</td><td>C</td></tr><tr><td>D</td><td>E</td><td>F</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    expect(findAll(tbl, 'w:gridCol')).toHaveLength(3);
    assertColumnCountConsistent(tbl);

    const row0Cells = findAll(findAll(tbl, 'w:tr')[0], 'w:tc');
    expect(row0Cells).toHaveLength(2);
    expect(Number(attr(find(row0Cells[0], 'w:gridSpan'), 'w:val'))).toBe(2);
  });

  it('W6 — rowspan emits vMerge restart on the origin cell and a bare vMerge continuation placeholder', async () => {
    const html = '<table><tr><td rowspan="2">A</td><td>B</td></tr><tr><td>C</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    expect(findAll(tbl, 'w:gridCol')).toHaveLength(2);
    assertColumnCountConsistent(tbl);

    const rows = findAll(tbl, 'w:tr');
    const row0Cells = findAll(rows[0], 'w:tc');
    const row1Cells = findAll(rows[1], 'w:tc');
    expect(row0Cells).toHaveLength(2);
    expect(row1Cells).toHaveLength(2); // continuation placeholder + C — not 1

    expect(attr(find(row0Cells[0], 'w:vMerge'), 'w:val')).toBe('restart');
    const continuationVMerge = find(row1Cells[0], 'w:vMerge');
    expect(continuationVMerge).toBeTruthy();
    expect(attr(continuationVMerge, 'w:val')).toBeUndefined(); // bare <w:vMerge/> = continue
    expect(textContent(row1Cells[0])).toBe(''); // continuation cell carries no duplicated content
    expect(textContent(row1Cells[1])).toBe('C');
  });

  it('W6 — combined rowspan+colspan: the continuation placeholder carries the same gridSpan as the origin', async () => {
    const html = '<table><tr><td rowspan="2" colspan="2">A</td><td>B</td></tr><tr><td>C</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    expect(findAll(tbl, 'w:gridCol')).toHaveLength(3);
    assertColumnCountConsistent(tbl);

    const rows = findAll(tbl, 'w:tr');
    const row1Cells = findAll(rows[1], 'w:tc');
    expect(row1Cells).toHaveLength(2); // continuation (gridSpan=2) + C
    expect(Number(attr(find(row1Cells[0], 'w:gridSpan'), 'w:val'))).toBe(2);
    expect(find(row1Cells[0], 'w:vMerge')).toBeTruthy();
  });

  it('W6 — a cell with multiple paragraphs renders as multiple w:p inside one w:tc', async () => {
    const html = '<table><tr><td><p>One</p><p>Two</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tc = find(xml['word/document.xml'], 'w:tc');
    expect(findAll(tc, 'w:p').map(textContent)).toEqual(['One', 'Two']);
  });

  it('W6 — a cell containing a nested table renders a real nested w:tbl, not flattened text', async () => {
    const html = '<table><tr><td><table><tr><td>Inner</td></tr></table></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbls = findAll(xml['word/document.xml'], 'w:tbl');
    expect(tbls).toHaveLength(2); // outer + nested
    const outerTc = find(xml['word/document.xml'], 'w:tc');
    expect(find(outerTc, 'w:tbl')).toBeTruthy();
    expect(textContent(outerTc)).toContain('Inner');
  });

  it('W7 — a link becomes w:hyperlink with a matching External relationship in document.xml.rels', async () => {
    const { buffer } = await convert('<p><a href="https://example.com">click</a></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const hyperlink = find(xml['word/document.xml'], 'w:hyperlink');
    expect(textContent(hyperlink)).toBe('click');
    const rId = attr(hyperlink, 'r:id');
    expect(rId).toBeTruthy();

    const rel = findAll(xml['word/_rels/document.xml.rels'], 'Relationship').find((r) => attr(r, 'Id') === rId);
    expect(rel).toBeTruthy();
    expect(attr(rel, 'Target')).toBe('https://example.com');
    expect(attr(rel, 'TargetMode')).toBe('External');
  });

  it('W7 — two different links get two distinct relationship ids, both resolvable', async () => {
    const html = '<p><a href="https://a.example">A</a> <a href="https://b.example">B</a></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const hyperlinks = findAll(xml['word/document.xml'], 'w:hyperlink');
    expect(hyperlinks).toHaveLength(2);
    const rels = findAll(xml['word/_rels/document.xml.rels'], 'Relationship');
    for (const link of hyperlinks) {
      const rId = attr(link, 'r:id');
      const rel = rels.find((r) => attr(r, 'Id') === rId);
      expect(rel).toBeTruthy();
    }
    const targets = hyperlinks.map((link) => {
      const rId = attr(link, 'r:id');
      return attr(rels.find((r) => attr(r, 'Id') === rId), 'Target');
    });
    expect(targets).toEqual(['https://a.example', 'https://b.example']);
  });

  it('an internal anchor link (href="#id") uses w:anchor, not a broken External relationship', async () => {
    const html = '<footer id="end">Footer</footer><p><a href="#end">Jump</a></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const link = find(xml['word/document.xml'], 'w:hyperlink');
    expect(attr(link, 'w:anchor')).toBe('end');
    expect(attr(link, 'r:id')).toBeUndefined();
    const rels = findAll(xml['word/_rels/document.xml.rels'], 'Relationship');
    expect(rels.some((r) => attr(r, 'Target') === '#end')).toBe(false);
  });

  it('the element with the matching id gets a real w:bookmarkStart/w:bookmarkEnd pair', async () => {
    const html = '<footer id="end">Footer</footer><p><a href="#end">Jump</a></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const bookmarkStart = find(xml['word/document.xml'], 'w:bookmarkStart');
    const bookmarkEnd = find(xml['word/document.xml'], 'w:bookmarkEnd');
    expect(attr(bookmarkStart, 'w:name')).toBe('end');
    expect(attr(bookmarkEnd, 'w:id')).toBe(attr(bookmarkStart, 'w:id'));
  });

  it('an id containing characters invalid in an OOXML bookmark name is sanitized consistently on both ends', async () => {
    const html = '<footer id="my-section.2">Footer</footer><p><a href="#my-section.2">Jump</a></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const link = find(xml['word/document.xml'], 'w:hyperlink');
    const bookmarkStart = find(xml['word/document.xml'], 'w:bookmarkStart');
    expect(attr(link, 'w:anchor')).toBe(attr(bookmarkStart, 'w:name'));
    expect(attr(bookmarkStart, 'w:name')).toMatch(/^[A-Za-z_][A-Za-z0-9_]*$/);
  });

  it('a real external link is unaffected by the internal-anchor handling (no regression)', async () => {
    const { buffer } = await convert('<p><a href="https://example.com">Ext</a></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const link = find(xml['word/document.xml'], 'w:hyperlink');
    expect(attr(link, 'r:id')).toBeTruthy();
    expect(attr(link, 'w:anchor')).toBeUndefined();
  });

  it('a <div class="pagebreak"> styled with page-break-before:always renders as a real page break', async () => {
    const html = '<p>Before</p><div style="page-break-before:always">After</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const pageBreak = find(xml['word/document.xml'], 'w:br');
    expect(attr(pageBreak, 'w:type')).toBe('page');
  });

  it('an unchecked checkbox renders the hollow glyph, a checked one renders the filled glyph', async () => {
    const html = '<p><input type="checkbox"> A</p><p><input type="checkbox" checked> B</p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const text = textContent(xml['word/document.xml']);
    expect(text).toContain('☐ A');
    expect(text).toContain('☑ B');
  });

  it('a <select> renders only its selected option\'s text, not every option run together', async () => {
    const html = '<p><select><option>One</option><option selected>Two</option></select></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(textContent(paragraphs[0])).toBe('Two');
  });

  it('padding on a bordered <div> becomes the pBdr w:space (distance from border to text), not the default 1pt', async () => {
    const html = '<div style="border:2px dashed red;padding:10px">Dashed</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const top = find(find(find(xml['word/document.xml'], 'w:p'), 'w:pBdr'), 'w:top');
    expect(Number(attr(top, 'w:space'))).toBe(8); // 10px = 7.5pt, rounded to 8
  });

  it('padding on a background-only <div> (no border) synthesizes an invisible same-color border, so top/bottom padding isn\'t dropped', async () => {
    const html = '<div style="background:linear-gradient(90deg,#ffd,#cdf);padding:10px">Gradient background</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    const pBdr = find(p, 'w:pBdr');
    expect(pBdr).toBeTruthy();
    const top = find(pBdr, 'w:top');
    expect(attr(top, 'w:color')).toBe('FFFFDD'); // matches the resolved background fill, so it renders invisibly
    expect(Number(attr(top, 'w:space'))).toBe(8); // 10px = 7.5pt, rounded to 8 — vertical padding now actually renders
    expect(find(p, 'w:ind')).toBeFalsy(); // superseded by pBdr's w:space, not double-applied
  });

  it('padding with neither a border nor a background still falls back to a horizontal w:ind', async () => {
    const html = '<div style="padding:10px">Plain padding</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(find(p, 'w:pBdr')).toBeFalsy();
    const ind = find(p, 'w:ind');
    expect(Number(attr(ind, 'w:left'))).toBe(150); // 10px = 7.5pt * 20 twips/pt
    expect(Number(attr(ind, 'w:right'))).toBe(150);
  });

  it('a bordered <div> with no explicit padding falls back to the small fixed default border-to-text space', async () => {
    const { buffer } = await convert('<div style="border:2px solid #444">Boxed</div>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const top = find(find(find(xml['word/document.xml'], 'w:p'), 'w:pBdr'), 'w:top');
    expect(Number(attr(top, 'w:space'))).toBe(1);
  });

  it('a <div> with no border and no padding gets no w:ind (no regression)', async () => {
    const { buffer } = await convert('<div>Plain</div>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(find(p, 'w:ind')).toBeFalsy();
  });

  it('a <div style="border-left:...;background:...;padding:..."> renders a real left accent stripe plus invisible same-color borders on the other 3 sides for padding', async () => {
    const html = '<div style="border-left:5px solid #1f4e79;background:#eef6ff;padding:12px">Note</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const pBdr = find(xml['word/document.xml'], 'w:pBdr');
    const left = find(pBdr, 'w:left');
    expect(attr(left, 'w:color')).toBe('1F4E79');
    expect(Number(attr(left, 'w:sz'))).toBe(30); // 5px border-left, a real visible stripe
    const top = find(pBdr, 'w:top');
    expect(attr(top, 'w:color')).toBe('EEF6FF'); // matches the background fill, so it renders invisibly
    expect(Number(attr(top, 'w:sz'))).toBe(2); // hairline, not the 5px accent width
    expect(Number(attr(top, 'w:space'))).toBe(9); // 12px = 9pt padding, preserved on every side
  });

  it('a <div style="border-left:..."> with no background/padding renders only the left edge, no invisible fallback on the others', async () => {
    const html = '<div style="border-left:5px solid #1f4e79">Note</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const pBdr = find(xml['word/document.xml'], 'w:pBdr');
    expect(find(pBdr, 'w:left')).toBeTruthy();
    expect(find(pBdr, 'w:top')).toBeFalsy();
    expect(find(pBdr, 'w:right')).toBeFalsy();
    expect(find(pBdr, 'w:bottom')).toBeFalsy();
  });

  const TINY_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

  it('W8 — an image is actually embedded (real media file + relationship + w:blip), not a text placeholder', async () => {
    const { buffer } = await convert(`<img src="${TINY_PNG}" alt="Dot">`, { format: 'docx' });
    const { zip, xml } = readDocxParts(buffer);

    // Regression guard: this is the exact bug from the prior attempt — always
    // falling back to literal placeholder text instead of a real drawing.
    expect(textContent(xml['word/document.xml'])).not.toContain('[Image:');

    const blip = find(xml['word/document.xml'], 'a:blip');
    expect(blip).toBeTruthy();
    const rId = attr(blip, 'r:embed');
    expect(rId).toBeTruthy();

    const rel = findAll(xml['word/_rels/document.xml.rels'], 'Relationship').find((r) => attr(r, 'Id') === rId);
    expect(rel).toBeTruthy();
    expect(attr(rel, 'Target')).toBe('media/image1.png');
    expect(zip.has('word/media/image1.png')).toBe(true);
    expect(zip.read('word/media/image1.png').length).toBeGreaterThan(0);
  });

  it('W8 — [Content_Types].xml declares a Default for the embedded image extension', async () => {
    const { buffer } = await convert(`<img src="${TINY_PNG}">`, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const defaults = findAll(xml['[Content_Types].xml'], 'Default').map((el) => attr(el, 'Extension'));
    expect(defaults).toContain('png');
  });

  it('W8 — a figure renders the image plus a caption paragraph', async () => {
    const html = `<figure><img src="${TINY_PNG}" alt="Dot"><figcaption>A tiny dot</figcaption></figure>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    expect(find(xml['word/document.xml'], 'a:blip')).toBeTruthy();
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(paragraphs.some((p) => textContent(p) === 'A tiny dot')).toBe(true);
  });

  const TINY_GIF = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7';

  it('W8 — two distinct images get distinct relationship ids and media filenames', async () => {
    const html = `<img src="${TINY_PNG}"><img src="${TINY_GIF}">`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { zip, xml } = readDocxParts(buffer);
    const blips = findAll(xml['word/document.xml'], 'a:blip');
    expect(blips).toHaveLength(2);
    const rIds = blips.map((b) => attr(b, 'r:embed'));
    expect(new Set(rIds).size).toBe(2);
    expect(zip.has('word/media/image1.png')).toBe(true);
    expect(zip.has('word/media/image2.gif')).toBe(true);
  });

  it('W8 — the same image src used twice is embedded once and both references share the relationship', async () => {
    const html = `<img src="${TINY_PNG}"><img src="${TINY_PNG}">`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { zip, xml } = readDocxParts(buffer);
    const blips = findAll(xml['word/document.xml'], 'a:blip');
    expect(blips).toHaveLength(2);
    expect(attr(blips[0], 'r:embed')).toBe(attr(blips[1], 'r:embed')); // deduped by resolveAssets
    expect(zip.has('word/media/image1.png')).toBe(true);
    expect(zip.has('word/media/image2.png')).toBe(false); // only one file was actually needed
  });

  it('an <svg> wrapping an embedded base64 image is rendered as a real embedded image, not dropped', async () => {
    const html = `<svg width="100" height="50"><image xlink:href="${TINY_PNG}"/></svg>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { zip, xml } = readDocxParts(buffer);
    expect(find(xml['word/document.xml'], 'a:blip')).toBeTruthy();
    expect(zip.has('word/media/image1.png')).toBe(true);
  });

  describe('SVG vector shapes', () => {
    it('a rect+circle logo draws two separately-anchored shapes, the circle centered within the rect', async () => {
      const html = '<svg width="200" height="120"><rect width="200" height="120" fill="#2563eb"/><circle cx="100" cy="60" r="35" fill="orange"/></svg>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const anchors = findAll(xml['word/document.xml'], 'wp:anchor');
      expect(anchors).toHaveLength(2);

      const [rectAnchor, circleAnchor] = anchors;
      const rectShape = find(rectAnchor, 'wps:wsp');
      const circleShape = find(circleAnchor, 'wps:wsp');
      expect(attr(find(rectShape, 'a:prstGeom'), 'prst')).toBe('rect');
      expect(attr(find(rectShape, 'a:srgbClr'), 'val')).toBe('2563EB');
      expect(attr(find(circleShape, 'a:prstGeom'), 'prst')).toBe('ellipse');
      expect(attr(find(circleShape, 'a:srgbClr'), 'val')).toBe('FFA500');

      // Both anchors share the same paragraph/column origin, so their own
      // posOffset + half-extent should land on the same center point.
      const posOffset = (anchor, tag) => Number(textContent(find(find(anchor, tag), 'wp:posOffset')));
      const extent = (anchor) => find(anchor, 'wp:extent');
      const rectCenterX = posOffset(rectAnchor, 'wp:positionH') + Number(attr(extent(rectAnchor), 'cx')) / 2;
      const circleCenterX = posOffset(circleAnchor, 'wp:positionH') + Number(attr(extent(circleAnchor), 'cx')) / 2;
      expect(circleCenterX).toBeCloseTo(rectCenterX, 0);
      const rectCenterY = posOffset(rectAnchor, 'wp:positionV') + Number(attr(extent(rectAnchor), 'cy')) / 2;
      const circleCenterY = posOffset(circleAnchor, 'wp:positionV') + Number(attr(extent(circleAnchor), 'cy')) / 2;
      expect(circleCenterY).toBeCloseTo(rectCenterY, 0);
    });

    it('the host paragraph reserves vertical space (w:spacing w:after) matching the SVG\'s own height', async () => {
      const html = '<svg width="200" height="120"><rect width="200" height="120" fill="red"/></svg>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const spacing = find(xml['word/document.xml'], 'w:spacing');
      // 120px * 9525 EMU/px / 12700 EMU/pt * 20 twips/pt = 1800 twips.
      expect(Number(attr(spacing, 'w:after'))).toBe(1800);
    });

    it('a polygon renders as a real a:custGeom closed path with the right fill/stroke', async () => {
      const html = '<svg width="100" height="100"><polygon points="0,0 10,0 5,10" fill="red" stroke="blue"/></svg>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const custGeom = find(xml['word/document.xml'], 'a:custGeom');
      expect(custGeom).toBeTruthy();
      expect(findAll(custGeom, 'a:lnTo')).toHaveLength(2);
      expect(find(custGeom, 'a:close')).toBeTruthy();
      const fills = findAll(xml['word/document.xml'], 'a:srgbClr').map((el) => attr(el, 'val'));
      expect(fills).toEqual(['FF0000', '0000FF']);
    });

    it('a line renders as a stroked, unfilled line-preset shape', async () => {
      const html = '<svg width="100" height="100"><line x1="0" y1="0" x2="100" y2="100" stroke="black"/></svg>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const shape = find(xml['word/document.xml'], 'wps:wsp');
      expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('line');
      expect(find(shape, 'a:ln')).toBeTruthy();
    });

    it('an SVG with only unsupported content (e.g. a <path>) still drops silently', async () => {
      const html = '<svg><path d="M10 10 L90 90"/></svg><p>after</p>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      expect(find(xml['word/document.xml'], 'wp:anchor')).toBeFalsy();
      expect(textContent(xml['word/document.xml'])).toContain('after');
    });
  });

  it('a <table style="border:..."> renders an explicit w:tblBorders with the parsed color/size, overriding the layout-table heuristic', async () => {
    const html = '<table style="border:2px solid #2563eb"><tr><td><p>Card content</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    const top = find(find(tbl, 'w:tblBorders'), 'w:top');
    expect(attr(top, 'w:val')).toBe('single');
    expect(attr(top, 'w:color')).toBe('2563EB');
    expect(Number(attr(top, 'w:sz'))).toBe(12); // 2px * 0.75 = 1.5pt * 8 = 12 eighths-of-a-point
  });

  it('a <table style="border:..."> does not draw inner row/column dividers — only the outer perimeter, matching CSS border semantics', async () => {
    const html =
      '<table style="border:1px solid #e5e7eb"><tr><td>Top strip</td></tr><tr><td><p>Content</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const borders = find(xml['word/document.xml'], 'w:tblBorders');
    expect(attr(find(borders, 'w:insideH'), 'w:val')).toBe('none');
    expect(attr(find(borders, 'w:insideV'), 'w:val')).toBe('none');
    expect(attr(find(borders, 'w:top'), 'w:val')).toBe('single');
  });

  it('an image\'s real intrinsic size wins over a declared width/height attribute, not the other way around', async () => {
    // TINY_PNG's own real decoded size is 1x1 — a stale/wrong width="209"
    // height="56" attribute (the common real-world case: markup authored
    // against a placeholder, later swapped for a different file) must NOT
    // override what the file actually is.
    const html = `<img src="${TINY_PNG}" width="209" height="56" alt="Logo">`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const extent = find(xml['word/document.xml'], 'wp:extent');
    expect(Number(attr(extent, 'cx'))).toBe(9525); // 1px * 9525 EMU/px, not 209px's 1990725
    expect(Number(attr(extent, 'cy'))).toBe(9525);
  });

  it('an image wider than the page content width is scaled down, preserving aspect ratio', async () => {
    const bigPng = `data:image/png;base64,${buildTestPng(2000, 1000).toString('base64')}`;
    const html = `<img src="${bigPng}" alt="Big">`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const extent = find(xml['word/document.xml'], 'wp:extent');
    const cx = Number(attr(extent, 'cx'));
    const cy = Number(attr(extent, 'cy'));
    expect(cx).toBeLessThanOrEqual(9360 * 635); // CONTENT_WIDTH_TWIPS in EMU
    expect(Math.abs(cx / cy - 2)).toBeLessThan(0.01); // 2000:1000 aspect ratio preserved
  });

  it('an image with no declared width/height uses its own real intrinsic size (read from the bytes during asset resolution), not a fixed placeholder', async () => {
    const { buffer } = await convert(`<img src="${TINY_PNG}" alt="Dot">`, { format: 'docx' }); // a real 1x1 PNG
    const { xml } = readDocxParts(buffer);
    const extent = find(xml['word/document.xml'], 'wp:extent');
    expect(Number(attr(extent, 'cx'))).toBe(9525); // 1px * 9525 EMU/px
    expect(Number(attr(extent, 'cy'))).toBe(9525);
  });

  it('an image with no declared width/height AND no decodable intrinsic size (e.g. GIF) falls back to the generic placeholder', async () => {
    const { buffer } = await convert(`<img src="${TINY_GIF}" alt="Dot">`, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const extent = find(xml['word/document.xml'], 'wp:extent');
    expect(Number(attr(extent, 'cx'))).toBe(4572000);
    expect(Number(attr(extent, 'cy'))).toBe(2743200);
  });

  it('<td width="%"> proportions become real w:gridCol widths instead of an equal split', async () => {
    const html = '<table><tr><td width="52%"><p>Wide</p></td><td width="32%"><p>Mid</p></td><td width="16%"><p>Narrow</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    const cols = findAll(tbl, 'w:gridCol').map((c) => Number(attr(c, 'w:w')));
    const total = cols.reduce((a, w) => a + w, 0);
    expect(cols.map((w) => Math.round((w / total) * 100))).toEqual([52, 32, 16]);
    expect(attr(find(tbl, 'w:tblLayout'), 'w:type')).toBe('fixed');
  });

  it('a table with no declared column widths but equal-length content still splits evenly (no regression)', async () => {
    const html = '<table><tr><td>A</td><td>B</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    const cols = findAll(tbl, 'w:gridCol').map((c) => Number(attr(c, 'w:w')));
    expect(cols[0]).toBe(cols[1]);
  });

  it('a table with no declared column widths sizes each column by its longest content, not a flat equal split', async () => {
    // Mirrors a real invoice layout: a long label column next to a short
    // single-word "Subtotal"/"Amount"-style column — an equal split badly
    // squeezes the short column even though it never needs to wrap.
    const html = '<table>' +
      '<tr><td>A very long description that runs on for quite a while</td><td>Amt</td></tr>' +
      '<tr><td>Short</td><td>$1</td></tr>' +
      '</table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    const cols = findAll(tbl, 'w:gridCol').map((c) => Number(attr(c, 'w:w')));
    expect(cols[0]).toBeGreaterThan(cols[1] * 2);
    expect(cols[0] + cols[1]).toBe(9360); // still sums to the full available content width (fixed layout)
  });

  it('a colspan\'d cell\'s content length is divided across the columns it spans, not applied in full to each', async () => {
    const html = '<table>' +
      '<tr><td colspan="2">' + 'x'.repeat(40) + '</td></tr>' +
      '<tr><td>A</td><td>B</td></tr>' +
      '</table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    const cols = findAll(tbl, 'w:gridCol').map((c) => Number(attr(c, 'w:w')));
    expect(cols[0]).toBe(cols[1]); // still split evenly between the two, not each blown out to 40 chars
  });

  it('a pure spacer column (empty in every row) still gets a nonzero minimum width, not ~0', async () => {
    const html = '<table>' +
      '<tr><td>A very long description that runs on for quite a while indeed</td><td></td></tr>' +
      '<tr><td>Short</td><td></td></tr>' +
      '</table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    const cols = findAll(tbl, 'w:gridCol').map((c) => Number(attr(c, 'w:w')));
    expect(cols[1]).toBeGreaterThan(0);
  });

  it('a large font-size cell (e.g. a title banner) spread across a colspan gets enough combined column width to fit', async () => {
    const html = '<table><tr><td colspan="3" style="font-size:60px;font-weight:bold">INVOICE</td></tr>' +
      '<tr><td>x</td><td>x</td><td>x</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbl = find(xml['word/document.xml'], 'w:tbl');
    const cols = findAll(tbl, 'w:gridCol').map((c) => Number(attr(c, 'w:w')));
    const span = cols[0] + cols[1] + cols[2];
    // 7 chars * (45pt/11pt) * 140 twips/char ≈ 4000 twips minimum required.
    expect(span).toBeGreaterThan(3900);
  });

  it('a nested table inherits its containing cell\'s width instead of claiming the full page width', async () => {
    const html =
      '<table><tr><td width="80%"><p>Wide</p></td>' +
      '<td width="20%"><table><tr><td>Nested</td></tr></table></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tbls = findAll(xml['word/document.xml'], 'w:tbl');
    expect(tbls).toHaveLength(2);
    const outerCol20 = findAll(tbls[0], 'w:gridCol').map((c) => Number(attr(c, 'w:w')))[1];
    const nestedCol = Number(attr(find(tbls[1], 'w:gridCol'), 'w:w'));
    expect(nestedCol).toBe(outerCol20); // nested table matches its cell's width, not the full page
  });

  it('a decorative <td style="height:4px"> row gets hRule="exact" so it actually shrinks to that height', async () => {
    const html = '<table><tr><td style="height:4px">&nbsp;</td></tr><tr><td><p>Content</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const rows = findAll(xml['word/document.xml'], 'w:tr');
    const trHeight = find(rows[0], 'w:trHeight');
    expect(attr(trHeight, 'w:hRule')).toBe('exact');
    expect(Number(attr(trHeight, 'w:val'))).toBe(60); // 4px * 15 twips/px
    expect(find(rows[1], 'w:trHeight')).toBeFalsy(); // no declared height on the content row
  });

  it('a row with real content and an explicit height uses hRule="atLeast" (no clipping)', async () => {
    const html = '<table><tr><td style="height:60px"><p>Real content</p></td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const trHeight = find(xml['word/document.xml'], 'w:trHeight');
    expect(attr(trHeight, 'w:hRule')).toBe('atLeast');
  });

  it('a <td style="border-bottom:..."> renders as a w:tcBorders bottom edge, independent of any table-level border', async () => {
    const html = '<table><tr><td style="border-bottom:1px solid #e5e7eb">&nbsp;</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tc = find(xml['word/document.xml'], 'w:tc');
    const borders = find(tc, 'w:tcBorders');
    const bottom = find(borders, 'w:bottom');
    expect(attr(bottom, 'w:val')).toBe('single');
    expect(attr(bottom, 'w:color')).toBe('E5E7EB');
    expect(find(borders, 'w:top')).toBeFalsy();
  });

  it('a <td style="border:none"> in a data table (which gets a TableGrid style) explicitly suppresses the border, rather than inheriting the table style\'s default grid line', async () => {
    // A real-world pattern: a base `td{border:1px solid #ccc}` tag rule plus a
    // `.no-border{border:none}` class on purely-positional cells within one
    // large layout+data table (e.g. an invoice template).
    const html = '<html><head><style>td{border:1px solid #ccc}.no-border{border:none}</style></head>' +
      '<body><table><tr><td>Data</td><td class="no-border">Layout</td></tr></table></body></html>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const cells = findAll(xml['word/document.xml'], 'w:tc');
    const dataBorders = find(cells[0], 'w:tcBorders');
    expect(attr(find(dataBorders, 'w:top'), 'w:val')).toBe('single');
    const layoutBorders = find(cells[1], 'w:tcBorders');
    expect(attr(find(layoutBorders, 'w:top'), 'w:val')).toBe('nil');
    expect(attr(find(layoutBorders, 'w:left'), 'w:val')).toBe('nil');
    expect(attr(find(layoutBorders, 'w:bottom'), 'w:val')).toBe('nil');
    expect(attr(find(layoutBorders, 'w:right'), 'w:val')).toBe('nil');
  });

  it('a <td> with no border style at all still gets no w:tcBorders (inherits the table style, no regression)', async () => {
    const html = '<table><tr><td>Plain</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tc = find(xml['word/document.xml'], 'w:tc');
    expect(find(tc, 'w:tcBorders')).toBeFalsy();
  });

  it('a <td style="padding:8px"> renders w:tcMar with the padding on all four sides', async () => {
    const html = '<table><tr><td style="padding:8px">A</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tc = find(xml['word/document.xml'], 'w:tc');
    const tcMar = find(tc, 'w:tcMar');
    expect(attr(find(tcMar, 'w:top'), 'w:w')).toBe('120'); // 8px = 6pt * 20 twips/pt
    expect(attr(find(tcMar, 'w:left'), 'w:w')).toBe('120');
  });

  it('a <td> with no padding style renders no w:tcMar (Word\'s own default cell margins apply)', async () => {
    const html = '<table><tr><td>A</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const tc = find(xml['word/document.xml'], 'w:tc');
    expect(find(tc, 'w:tcMar')).toBeFalsy();
  });

  it('a <th> renders center-aligned by default; a <td> in the same row does not', async () => {
    const html = '<table><tr><th>Head</th><td>Data</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(attr(find(paragraphs[0], 'w:jc'), 'w:val')).toBe('center');
    expect(find(paragraphs[1], 'w:jc')).toBeFalsy();
  });

  it('an explicit text-align on a <th> overrides its center default', async () => {
    const html = '<table><tr><th style="text-align:right">Head</th></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(attr(find(p, 'w:jc'), 'w:val')).toBe('right');
  });

  it('<u>/<ins> render w:u, <s>/<strike>/<del> render w:strike', async () => {
    const html = '<p><u>A</u> <ins>B</ins> <s>C</s> <strike>D</strike> <del>E</del></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const runs = findAll(xml['word/document.xml'], 'w:r').filter((r) => textContent(r).trim() !== '');
    const byText = (t) => runs.find((r) => textContent(r) === t);
    expect(find(byText('A'), 'w:u')).toBeTruthy();
    expect(find(byText('B'), 'w:u')).toBeTruthy();
    expect(find(byText('C'), 'w:strike')).toBeTruthy();
    expect(find(byText('D'), 'w:strike')).toBeTruthy();
    expect(find(byText('E'), 'w:strike')).toBeTruthy();
  });

  it('<mark> renders as w:shd yellow highlight', async () => {
    const { buffer } = await convert('<p><mark>Hi</mark></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'Hi');
    expect(attr(find(run, 'w:shd'), 'w:fill')).toBe('FFFF00');
  });

  it('<sub>/<sup> render w:vertAlign subscript/superscript', async () => {
    const { buffer } = await convert('<p>H<sub>2</sub>O x<sup>2</sup></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const runs = findAll(xml['word/document.xml'], 'w:r');
    const sub = runs.find((r) => textContent(r) === '2' && find(r, 'w:vertAlign') && attr(find(r, 'w:vertAlign'), 'w:val') === 'subscript');
    const sup = runs.find((r) => textContent(r) === '2' && find(r, 'w:vertAlign') && attr(find(r, 'w:vertAlign'), 'w:val') === 'superscript');
    expect(sub).toBeTruthy();
    expect(sup).toBeTruthy();
  });

  it('<kbd> gets the same Courier New + shading treatment as <code>', async () => {
    const { buffer } = await convert('<p><kbd>Ctrl+C</kbd></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'Ctrl+C');
    expect(attr(find(run, 'w:rFonts'), 'w:ascii')).toBe('Courier New');
  });

  it('<var>/<cite> render italic', async () => {
    const { buffer } = await convert('<p><var>x</var> <cite>Book</cite></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const runs = findAll(xml['word/document.xml'], 'w:r');
    expect(find(runs.find((r) => textContent(r) === 'x'), 'w:i')).toBeTruthy();
    expect(find(runs.find((r) => textContent(r) === 'Book'), 'w:i')).toBeTruthy();
  });

  it('<q> renders its content surrounded by literal quotation marks', async () => {
    const { buffer } = await convert('<p><q>Quote</q></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const paragraphs = findAll(xml['word/document.xml'], 'w:p');
    expect(textContent(paragraphs[0])).toBe('"Quote"');
  });

  it('a <span style="background:yellow"> (shorthand) renders the same highlight as background-color', async () => {
    const { buffer } = await convert('<p><span style="background:yellow">Hi</span></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'Hi');
    expect(attr(find(run, 'w:shd'), 'w:fill')).toBe('FFFF00');
  });

  it('<abbr> keeps its visible text even though the tooltip title has no DOCX equivalent', async () => {
    const { buffer } = await convert('<p><abbr title="HyperText Markup Language">HTML</abbr></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    expect(textContent(xml['word/document.xml'])).toContain('HTML');
  });

  it('font-family:Arial renders as w:rFonts on the run', async () => {
    const { buffer } = await convert('<p><span style="font-family:Arial">Arial</span></p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'Arial');
    expect(attr(find(run, 'w:rFonts'), 'w:ascii')).toBe('Arial');
    expect(attr(find(run, 'w:rFonts'), 'w:hAnsi')).toBe('Arial');
  });

  it('a quoted, comma-separated font-family stack uses only the first font name', async () => {
    const html = '<p><span style="font-family:\'Times New Roman\', serif">Times</span></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'Times');
    expect(attr(find(run, 'w:rFonts'), 'w:ascii')).toBe('Times New Roman');
  });

  it('an explicit font-family on a <code> span overrides the default Courier New', async () => {
    const html = '<p><code style="font-family:Consolas">x()</code></p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'x()');
    expect(attr(find(run, 'w:rFonts'), 'w:ascii')).toBe('Consolas');
  });

  it.each([
    ['xx-small', 6.75], ['small', 9.75], ['medium', 12], ['large', 13.5], ['xx-large', 24],
  ])('font-size:%s (CSS absolute keyword) renders at its fixed size (%spt)', async (keyword, pt) => {
    const { buffer } = await convert(`<p><span style="font-size:${keyword}">x</span></p>`, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const run = findAll(xml['word/document.xml'], 'w:r').find((r) => textContent(r) === 'x');
    expect(Number(attr(find(run, 'w:sz'), 'w:val'))).toBe(Math.round(pt * 2));
  });

  it('a <div class="box"> with a border and background renders as a real bordered/shaded paragraph, not plain text', async () => {
    const html =
      '<html><head><style>.box{border:2px solid #444;background-color:#eee}</style></head>' +
      '<body><div class="box">Boxed</div></body></html>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = findAll(xml['word/document.xml'], 'w:p').find((p) => textContent(p) === 'Boxed');
    const pPr = find(p, 'w:pPr');
    const top = find(find(pPr, 'w:pBdr'), 'w:top');
    expect(attr(top, 'w:val')).toBe('single');
    expect(attr(top, 'w:color')).toBe('444444');
    expect(attr(find(pPr, 'w:shd'), 'w:fill')).toBe('EEEEEE');
  });

  it.each([
    ['dashed', 'red', 'dashed', 'FF0000'],
    ['dotted', 'blue', 'dotted', '0000FF'],
    ['double', 'green', 'double', '008000'],
  ])('a <div> border-style:%s maps to the matching OOXML w:val (not always "single")', async (cssStyle, cssColor, ooxmlVal, color) => {
    const html = `<div style="border:2px ${cssStyle} ${cssColor}">X</div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    const top = find(find(find(p, 'w:pPr'), 'w:pBdr'), 'w:top');
    expect(attr(top, 'w:val')).toBe(ooxmlVal);
    expect(attr(top, 'w:color')).toBe(color);
  });

  it.each(['dashed', 'dotted', 'double'])(
    'a <div> with border-style:%s but no color still maps the style correctly, defaulting to black',
    async (cssStyle) => {
      const { buffer } = await convert(`<div style="border-width:2px;border-style:${cssStyle}">X</div>`, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const top = find(find(find(find(xml['word/document.xml'], 'w:p'), 'w:pPr'), 'w:pBdr'), 'w:top');
      expect(attr(top, 'w:val')).toBe(cssStyle);
      expect(attr(top, 'w:color')).toBe('000000');
    },
  );

  it('all four sides of a <div> border are emitted identically from one CSS border declaration', async () => {
    const { buffer } = await convert('<div style="border:2px solid #444">X</div>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const pBdr = find(xml['word/document.xml'], 'w:pBdr');
    for (const edge of ['w:top', 'w:left', 'w:bottom', 'w:right']) {
      const el = find(pBdr, edge);
      expect(attr(el, 'w:val')).toBe('single');
      expect(attr(el, 'w:color')).toBe('444444');
    }
  });

  it('a <div> with an unrepresentable gradient background falls back to a flat fill of its first color stop', async () => {
    const html = '<div style="background:linear-gradient(90deg,#ffd,#cdf)">Gradient</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(textContent(p)).toBe('Gradient');
    const pPr = find(p, 'w:pPr');
    expect(attr(find(pPr, 'w:shd'), 'w:fill')).toBe('FFFFDD');
  });

  it('a gradient with no resolvable color stops (e.g. an image-only background) renders plain text, no fabricated color', async () => {
    const html = '<div style="background:linear-gradient(90deg, url(x.png))">Plain</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(textContent(p)).toBe('Plain');
    const pPr = find(p, 'w:pPr');
    expect(pPr ? find(pPr, 'w:shd') : undefined).toBeFalsy();
  });

  it('a <div> with no border/background style renders as a plain paragraph (no regression)', async () => {
    const { buffer } = await convert('<div>Plain</div>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    expect(textContent(p)).toBe('Plain');
    expect(find(p, 'w:pBdr')).toBeFalsy();
  });

  it('the Normal style has a non-zero default "space after" so plain paragraphs aren\'t crammed together', async () => {
    const { buffer } = await convert('<p>A</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const normalStyle = findAll(xml['word/styles.xml'], 'w:style').find((s) => attr(s, 'w:styleId') === 'Normal');
    expect(Number(attr(find(normalStyle, 'w:spacing'), 'w:after'))).toBeGreaterThan(0);
  });

  it.each([1, 2, 3, 4, 5, 6])('Heading%i has non-zero default spacing before and after', async (level) => {
    const { buffer } = await convert('<p>x</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const headingStyle = findAll(xml['word/styles.xml'], 'w:style').find((s) => attr(s, 'w:styleId') === `Heading${level}`);
    const spacing = find(headingStyle, 'w:spacing');
    expect(Number(attr(spacing, 'w:before'))).toBeGreaterThan(0);
    expect(Number(attr(spacing, 'w:after'))).toBeGreaterThan(0);
  });

  it('ListBullet/ListNumber explicitly zero out spacing so adjacent list items stay tightly packed', async () => {
    const { buffer } = await convert('<p>x</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    for (const styleId of ['ListBullet', 'ListNumber']) {
      const style = findAll(xml['word/styles.xml'], 'w:style').find((s) => attr(s, 'w:styleId') === styleId);
      const spacing = find(style, 'w:spacing');
      expect(Number(attr(spacing, 'w:before'))).toBe(0);
      expect(Number(attr(spacing, 'w:after'))).toBe(0);
    }
  });

  it('a <div style="margin:10px 0"> renders w:spacing before/after on that paragraph, overriding the style default', async () => {
    const { buffer } = await convert('<div style="margin:10px 0">Boxed</div>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    const spacing = find(p, 'w:spacing');
    expect(Number(attr(spacing, 'w:before'))).toBe(150); // 10px = 7.5pt * 20 twips/pt = 150
    expect(Number(attr(spacing, 'w:after'))).toBe(150);
  });

  it('a <div style="margin-bottom:...;"> only sets w:after, leaving w:before absent (inherits the style default)', async () => {
    const { buffer } = await convert('<div style="margin-bottom:20px">Boxed</div>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const p = find(xml['word/document.xml'], 'w:p');
    const spacing = find(p, 'w:spacing');
    expect(attr(spacing, 'w:before')).toBeUndefined();
    expect(Number(attr(spacing, 'w:after'))).toBe(300); // 20px = 15pt * 20 = 300
  });

  it('a fonts/sizes/weights showcase renders each distinctly (font family, absolute size keywords, and numeric weights)', async () => {
    const html = `
      <p style="font-family:Arial">Arial</p>
      <p style="font-family:'Times New Roman'">Times New Roman</p>
      <p style="font-family:Georgia">Georgia</p>
      <p style="font-size:small">Small</p>
      <p style="font-size:large">Large</p>
      <p style="font-weight:100">100</p>
      <p style="font-weight:400">400</p>
      <p style="font-weight:700">700</p>
      <p style="font-weight:900">900</p>
    `;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const runs = findAll(xml['word/document.xml'], 'w:r');
    const byText = (t) => runs.find((r) => textContent(r) === t);

    expect(attr(find(byText('Arial'), 'w:rFonts'), 'w:ascii')).toBe('Arial');
    expect(attr(find(byText('Times New Roman'), 'w:rFonts'), 'w:ascii')).toBe('Times New Roman');
    expect(attr(find(byText('Georgia'), 'w:rFonts'), 'w:ascii')).toBe('Georgia');

    expect(Number(attr(find(byText('Small'), 'w:sz'), 'w:val'))).toBe(Math.round(9.75 * 2));
    expect(Number(attr(find(byText('Large'), 'w:sz'), 'w:val'))).toBe(Math.round(13.5 * 2));

    expect(find(byText('100'), 'w:b')).toBeFalsy();
    expect(find(byText('400'), 'w:b')).toBeFalsy();
    expect(find(byText('700'), 'w:b')).toBeTruthy();
    expect(find(byText('900'), 'w:b')).toBeTruthy();
  });

  // The real-world failure mode this guards: sizes/weights declared once in
  // a <style> block and applied via class="..." (very common in CMS/email
  // HTML) rather than repeated inline on every element — previously silently
  // ignored entirely, since only the inline style="..." attribute was read.
  it('a sizes/weights showcase styled entirely via <style> classes (not inline style) still renders correctly', async () => {
    const html = `
      <html><head><style>
        .small{font-size:10px}.large{font-size:28px}
        .w100{font-weight:100}.w400{font-weight:400}.w700{font-weight:700}.w900{font-weight:900}
      </style></head><body>
      <p class="small">Small</p>
      <p>Normal</p>
      <p class="large">Large</p>
      <p class="w100">100</p><p class="w400">400</p><p class="w700">700</p><p class="w900">900</p>
      </body></html>
    `;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const runs = findAll(xml['word/document.xml'], 'w:r');
    const byText = (t) => runs.find((r) => textContent(r) === t);

    const smallSz = Number(attr(find(byText('Small'), 'w:sz'), 'w:val'));
    const largeSz = Number(attr(find(byText('Large'), 'w:sz'), 'w:val'));
    expect(find(byText('Normal'), 'w:sz')).toBeFalsy(); // no class, no explicit size
    expect(smallSz).toBeLessThan(largeSz); // the actual visible bug: these came out identical before this fix

    expect(find(byText('100'), 'w:b')).toBeFalsy();
    expect(find(byText('400'), 'w:b')).toBeFalsy();
    expect(find(byText('700'), 'w:b')).toBeTruthy();
    expect(find(byText('900'), 'w:b')).toBeTruthy();
  });

  it('a <td style="background-color:..."> renders w:shd on that cell only', async () => {
    const html = '<table><tr><td style="background-color:#2563eb">&nbsp;</td><td>plain</td></tr></table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const cells = findAll(xml['word/document.xml'], 'w:tc');
    expect(attr(find(cells[0], 'w:shd'), 'w:fill')).toBe('2563EB');
    expect(find(cells[1], 'w:shd')).toBeFalsy();
  });

  it('W9 — <title>/lang from the source HTML land in docProps/core.xml', async () => {
    const html = '<html lang="en"><head><title>My Doc</title></head><body><p>x</p></body></html>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const core = xml['docProps/core.xml'];
    expect(textContent(find(core, 'dc:title'))).toBe('My Doc');
    expect(textContent(find(core, 'dc:language'))).toBe('en');
  });

  it('W9 — convert() options.title/author override and populate docProps/core.xml', async () => {
    const { buffer } = await convert('<p>x</p>', { format: 'docx', title: 'Report', author: 'Jane Doe' });
    const { xml } = readDocxParts(buffer);
    const core = xml['docProps/core.xml'];
    expect(textContent(find(core, 'dc:title'))).toBe('Report');
    expect(textContent(find(core, 'dc:creator'))).toBe('Jane Doe');
  });

  it('W9 — no metadata present means no empty dc:title/dc:creator elements are fabricated', async () => {
    const { buffer } = await convert('<p>x</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const core = xml['docProps/core.xml'];
    expect(find(core, 'dc:title')).toBeUndefined();
    expect(find(core, 'dc:creator')).toBeUndefined();
  });

  it('W10 — defaults to US Letter portrait with 1in margins when no page options are given', async () => {
    const { buffer } = await convert('<p>x</p>', { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const pgSz = find(xml['word/document.xml'], 'w:pgSz');
    expect(Number(attr(pgSz, 'w:w'))).toBe(12240);
    expect(Number(attr(pgSz, 'w:h'))).toBe(15840);
    const pgMar = find(xml['word/document.xml'], 'w:pgMar');
    expect(Number(attr(pgMar, 'w:top'))).toBe(1440);
  });

  it('W10 — pageSize: "A4" sets the corresponding twip dimensions', async () => {
    const { buffer } = await convert('<p>x</p>', { format: 'docx', pageSize: 'A4' });
    const { xml } = readDocxParts(buffer);
    const pgSz = find(xml['word/document.xml'], 'w:pgSz');
    expect(Number(attr(pgSz, 'w:w'))).toBe(11906);
    expect(Number(attr(pgSz, 'w:h'))).toBe(16838);
  });

  it('W10 — orientation: "landscape" swaps width/height and sets w:orient', async () => {
    const { buffer } = await convert('<p>x</p>', { format: 'docx', orientation: 'landscape' });
    const { xml } = readDocxParts(buffer);
    const pgSz = find(xml['word/document.xml'], 'w:pgSz');
    expect(Number(attr(pgSz, 'w:w'))).toBe(15840);
    expect(Number(attr(pgSz, 'w:h'))).toBe(12240);
    expect(attr(pgSz, 'w:orient')).toBe('landscape');
  });

  it('W10 — margins (mm) convert to twips', async () => {
    const { buffer } = await convert('<p>x</p>', { format: 'docx', margins: { top: 10, right: 10, bottom: 10, left: 10 } });
    const { xml } = readDocxParts(buffer);
    const pgMar = find(xml['word/document.xml'], 'w:pgMar');
    expect(Number(attr(pgMar, 'w:top'))).toBe(567);
    expect(Number(attr(pgMar, 'w:left'))).toBe(567);
  });

  it('a table\'s column widths scale with pageSize/orientation/margins, not a hardcoded page size', async () => {
    const html = '<table><tr><td>A</td><td>B</td></tr></table>';
    const portrait = await convert(html, { format: 'docx' });
    const landscape = await convert(html, { format: 'docx', pageSize: 'Legal', orientation: 'landscape', margins: { top: 5, right: 5, bottom: 5, left: 5 } });
    const sumWidths = (buffer) => {
      const { xml } = readDocxParts(buffer);
      return findAll(xml['word/document.xml'], 'w:gridCol').reduce((sum, c) => sum + Number(attr(c, 'w:w')), 0);
    };
    expect(sumWidths(landscape.buffer)).toBeGreaterThan(sumWidths(portrait.buffer) * 1.5);
  });

  it('a table with no declared column widths always produces integer w:w values, even after deficit-boosting and slack distribution', async () => {
    // A colspan across an odd number of columns forces a fractional
    // per-column deficit share partway through the computation.
    const html = '<table>' +
      '<tr><td colspan="3">' + 'x'.repeat(37) + '</td></tr>' +
      '<tr><td>A</td><td>BB</td><td>CCC</td></tr>' +
      '</table>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const widths = findAll(xml['word/document.xml'], 'w:gridCol').map((c) => attr(c, 'w:w'));
    for (const w of widths) expect(w).toMatch(/^\d+$/);
  });

  // Every relationship id document.xml actually uses (r:id on a hyperlink,
  // r:embed on a blip) must resolve in .rels, and vice versa isn't required,
  // but every declared media relationship's target file must really exist in
  // the zip. This is the check the prior DOCX attempt never had at any
  // granularity — a whole-document exercise across every milestone at once.
  function collectRIdReferences(el, out = []) {
    if (el.type === 'element') {
      for (const key of ['r:id', 'r:embed']) {
        if (key in el.attrs) out.push(el.attrs[key]);
      }
      for (const child of el.children) collectRIdReferences(child, out);
    }
    return out;
  }

  const INTEGRATION_FIXTURE = `
    <html lang="en"><head><title>Quarterly Report</title></head><body>
    <h1>Quarterly Report</h1>
    <p>Summary with <strong>bold</strong>, <em>italic</em>, and a <a href="https://example.com">link</a>.</p>
    <ul>
      <li>Point one</li>
      <li>Point two
        <ul><li>Nested point</li></ul>
      </li>
    </ul>
    <ol><li>Step one</li><li>Step two</li></ol>
    <table>
      <thead><tr><th>Region</th><th>Revenue</th></tr></thead>
      <tbody>
        <tr><td rowspan="2">Americas</td><td>$1M</td></tr>
        <tr><td>$2M</td></tr>
      </tbody>
    </table>
    <pre><code>const x = 1;</code></pre>
    <blockquote><p>A closing remark.</p></blockquote>
    <figure><img src="${TINY_PNG}" alt="Chart"><figcaption>Fig. 1</figcaption></figure>
    <hr>
    </body></html>
  `;

  it('integration — a fixture exercising every milestone: every part is well-formed and every relationship reference resolves', async () => {
    const { buffer } = await convert(INTEGRATION_FIXTURE, { format: 'docx', author: 'Finance Team' });
    const { zip, xml } = readDocxParts(buffer);

    for (const part of REQUIRED_PARTS) {
      expect(zip.has(part), `missing part: ${part}`).toBe(true);
      expect(xml[part], `did not parse as XML: ${part}`).toBeTruthy();
    }

    const usedRIds = collectRIdReferences(xml['word/document.xml']);
    expect(usedRIds.length).toBeGreaterThan(0); // sanity: the fixture actually exercises links/images

    const declaredRels = findAll(xml['word/_rels/document.xml.rels'], 'Relationship');
    const declaredIds = new Set(declaredRels.map((r) => attr(r, 'Id')));
    for (const rId of usedRIds) {
      expect(declaredIds.has(rId), `document.xml references ${rId} but it's missing from .rels`).toBe(true);
    }

    for (const rel of declaredRels) {
      const target = attr(rel, 'Target');
      if (target.startsWith('media/')) {
        expect(zip.has(`word/${target}`), `.rels declares ${target} but it's missing from the zip`).toBe(true);
      }
    }

    for (const tbl of findAll(xml['word/document.xml'], 'w:tbl')) {
      assertColumnCountConsistent(tbl);
    }

    expect(textContent(xml['word/document.xml'])).not.toContain('[Image:');
    expect(textContent(find(xml['docProps/core.xml'], 'dc:creator'))).toBe('Finance Team');
  });

  // Regression tests for the exact "real-world email/CMS table layout" bugs
  // found by comparing an actual generated .docx against its source HTML's
  // browser rendering: an oversized logo crushing its column, equal-split
  // columns ignoring declared %, a 4px decorative strip rendering as a thick
  // bar, a table-level border leaking onto internal row/cell dividers, and a
  // per-cell border-bottom divider line being dropped entirely.
  describe('Word styling regressions — real-world table layout bugs', () => {
    // findAll recurses through all descendants, so a raw findAll(tbl,
    // 'w:gridCol') on a table containing nested tables would also pick up
    // the nested tables' own gridCols. w:tblGrid's only children are
    // gridCol elements (never a nested table), so scoping through the
    // table's own w:tblGrid isolates just its direct columns.
    function directGridCols(tbl) {
      return findAll(find(tbl, 'w:tblGrid'), 'w:gridCol').map((c) => Number(attr(c, 'w:w')));
    }


    it('a 3-card stat-row layout (colored 4px top strip + bordered card + proportional widths) matches the source exactly', async () => {
      const card = (color, label, value) => `
        <table width="32%" style="border:1px solid #e5e7eb">
          <tr><td style="height:4px;background-color:${color}">&nbsp;</td></tr>
          <tr><td style="padding:24px"><p>${label}</p><p>${value}</p></td></tr>
        </table>
      `;
      const html = `
        <table>
          <tr>
            <td width="32%">${card('#2563eb', 'Total Changes', '06')}</td>
            <td width="2%">&nbsp;</td>
            <td width="32%">${card('#ef4444', 'Remove', '02')}</td>
            <td width="2%">&nbsp;</td>
            <td width="32%">${card('#22c55e', 'Add or Replace', '01')}</td>
          </tr>
        </table>
      `;
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const doc = xml['word/document.xml'];

      // Outer row proportions: 32/2/32/2/32, not an equal fifths split.
      const outerTbl = findAll(doc, 'w:tbl')[0];
      const outerCols = directGridCols(outerTbl);
      const outerTotal = outerCols.reduce((a, w) => a + w, 0);
      expect(outerCols.map((w) => Math.round((w / outerTotal) * 100))).toEqual([32, 2, 32, 2, 32]);

      // Each card's colored strip row is exactly 4px, not a full text line.
      const trHeights = findAll(doc, 'w:trHeight');
      expect(trHeights).toHaveLength(3);
      for (const h of trHeights) {
        expect(attr(h, 'w:hRule')).toBe('exact');
        expect(Number(attr(h, 'w:val'))).toBe(60); // 4px * 15 twips/px
      }

      // Each card's border is single top/left/bottom/right, with no
      // insideH/insideV line splitting the strip from its content. (The
      // outer table also emits a w:tblBorders block, but with val="none" —
      // filter to the real, explicit single borders only.)
      const cardBorders = findAll(doc, 'w:tblBorders').filter((b) => attr(find(b, 'w:top'), 'w:val') === 'single');
      expect(cardBorders).toHaveLength(3);
      for (const borders of cardBorders) {
        expect(attr(find(borders, 'w:top'), 'w:color')).toBe('E5E7EB');
        expect(attr(find(borders, 'w:insideH'), 'w:val')).toBe('none');
        expect(attr(find(borders, 'w:insideV'), 'w:val')).toBe('none');
      }

      // Each strip's own background color survives independently.
      const fills = findAll(doc, 'w:shd').map((s) => attr(s, 'w:fill'));
      expect(fills).toEqual(expect.arrayContaining(['2563EB', 'EF4444', '22C55E']));
    });

    it('an oversized image inside a percentage-based table column does not blow out that column\'s declared width', async () => {
      const bigPng = `data:image/png;base64,${buildTestPng(2000, 1000).toString('base64')}`;
      const html =
        `<table><tr>` +
        `<td width="20%"><img src="${bigPng}"></td>` +
        `<td width="80%"><p>Everything else</p></td>` +
        `</tr></table>`;
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const doc = xml['word/document.xml'];

      const cols = directGridCols(find(doc, 'w:tbl'));
      const total = cols.reduce((a, w) => a + w, 0);
      expect(cols.map((w) => Math.round((w / total) * 100))).toEqual([20, 80]);

      // The image itself must fit within its OWN column's real width — a
      // narrow 20% column on a 9360-twip page is ~1872 twips, nowhere near
      // the full page width. Clamping only to the page (the previous,
      // weaker assertion here) let a real, larger-than-CSS-declared image
      // silently overflow past its own cell into whatever sits next to it —
      // Word doesn't grow/clip an inline picture to its cell, so it visibly
      // bled over the boundary instead of erroring.
      const extent = find(doc, 'wp:extent');
      const cx = Number(attr(extent, 'cx'));
      const cy = Number(attr(extent, 'cy'));
      expect(cx).toBeLessThanOrEqual(cols[0] * 635); // cols[0] is in twips (dxa); wp:extent is in EMU
      expect(Math.abs(cx / cy - 2)).toBeLessThan(0.01);
    });

    it('a colspan cell with a declared width shares that width evenly across the columns it spans', async () => {
      const html =
        '<table><tr><td width="60%" colspan="2"><p>Wide</p></td><td width="40%"><p>Narrow</p></td></tr>' +
        '<tr><td>A</td><td>B</td><td>C</td></tr></table>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const tbl = find(xml['word/document.xml'], 'w:tbl');
      const cols = directGridCols(tbl);
      expect(cols).toHaveLength(3);
      const total = cols.reduce((a, w) => a + w, 0);
      // 60% split evenly across its 2 spanned columns (30% each), plus 40%.
      expect(cols.map((w) => Math.round((w / total) * 100))).toEqual([30, 30, 40]);
    });

    it('column-width percentages that don\'t sum to 100 are normalized to fill the full content width', async () => {
      const html = '<table><tr><td width="30%"><p>A</p></td><td width="30%"><p>B</p></td><td width="30%"><p>C</p></td></tr></table>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const tbl = find(xml['word/document.xml'], 'w:tbl');
      const cols = directGridCols(tbl);
      const total = cols.reduce((a, w) => a + w, 0);
      expect(total).toBeGreaterThan(9000); // fills ~the full 9360-twip content width, not 90% of it
      expect(cols[0]).toBe(cols[1]);
      expect(cols[1]).toBe(cols[2]);
    });

    it('a table nested three levels deep inherits each ancestor\'s actual width, not the full page width at any level', async () => {
      const html =
        '<table><tr>' +
        '<td width="50%"><p>Outer left</p></td>' +
        '<td width="50%"><table><tr>' +
        '<td width="40%"><p>Mid left</p></td>' +
        '<td width="60%"><table><tr><td><p>Innermost</p></td></tr></table></td>' +
        '</tr></table></td>' +
        '</tr></table>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const tbls = findAll(xml['word/document.xml'], 'w:tbl');
      expect(tbls).toHaveLength(3);

      const outerCols = directGridCols(tbls[0]);
      const midCols = directGridCols(tbls[1]);
      const innerCol = directGridCols(tbls[2])[0];

      // Mid-level table's total width matches the outer table's 50% cell.
      expect(midCols.reduce((a, w) => a + w, 0)).toBe(outerCols[1]);
      // Innermost table's single column matches the mid table's 60% cell.
      expect(innerCol).toBe(midCols[1]);
    });

    it('two decorative rows in the same table with different declared heights each get their own exact height', async () => {
      const html =
        '<table>' +
        '<tr><td style="height:46px">&nbsp;</td></tr>' +
        '<tr><td style="height:1px;border-bottom:1px solid #e5e7eb">&nbsp;</td></tr>' +
        '<tr><td><p>Real content</p></td></tr>' +
        '</table>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const rows = findAll(xml['word/document.xml'], 'w:tr');
      expect(Number(attr(find(rows[0], 'w:trHeight'), 'w:val'))).toBe(690); // 46px * 15
      expect(Number(attr(find(rows[1], 'w:trHeight'), 'w:val'))).toBe(15); // 1px * 15
      expect(find(rows[2], 'w:trHeight')).toBeFalsy();
    });

    it('a cell with both a background-color and a border renders both w:shd and w:tcBorders together', async () => {
      const html = '<table><tr><td style="background-color:#2563eb;border-bottom:1px solid #e5e7eb">&nbsp;</td></tr></table>';
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const tc = find(xml['word/document.xml'], 'w:tc');
      expect(attr(find(tc, 'w:shd'), 'w:fill')).toBe('2563EB');
      expect(attr(find(find(tc, 'w:tcBorders'), 'w:bottom'), 'w:color')).toBe('E5E7EB');
    });

    it('multiple images are each sized independently (no shared/leaked state) — a decodable one uses its own real size, an undecodable one keeps its declared attrs', async () => {
      const html = `<img src="${TINY_PNG}" width="209" height="56"><img src="${TINY_GIF}" width="100" height="100">`;
      const { buffer } = await convert(html, { format: 'docx' });
      const { xml } = readDocxParts(buffer);
      const extents = findAll(xml['word/document.xml'], 'wp:extent');
      expect(extents).toHaveLength(2);
      // TINY_PNG is a real, decodable 1x1 image — its own intrinsic size
      // wins over the (stale) declared 209x56 attribute.
      expect(Number(attr(extents[0], 'cx'))).toBe(9525);
      expect(Number(attr(extents[0], 'cy'))).toBe(9525);
      // TINY_GIF isn't decoded by this project (no npm dependency for GIF)
      // — its declared attrs are the only size information available, so
      // they're still honored as the documented fallback.
      expect(Number(attr(extents[1], 'cx'))).toBe(100 * 9525);
      expect(Number(attr(extents[1], 'cy'))).toBe(100 * 9525);
    });
  });
});

describe('DOCX exporter — flex/grid layout', () => {
  it('a display:flex container renders as one table row, one column per item, instead of stacked paragraphs', async () => {
    const html = `<div style="display:flex;gap:20px">
      <div style="background:#ef4444;color:white">Box 1</div>
      <div style="background:#10b981;color:white">Box 2</div>
      <div style="background:#f59e0b;color:white">Box 3</div>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const rows = findAll(doc, 'w:tr');
    expect(rows).toHaveLength(1);
    // item, spacer, item, spacer, item — see the gap test below for the
    // blank spacer cells woven in between real ones.
    const cells = findAll(rows[0], 'w:tc');
    expect(cells).toHaveLength(5);
    const itemCells = [cells[0], cells[2], cells[4]];
    expect(itemCells.map((tc) => textContent(find(tc, 'w:t')))).toEqual(['Box 1', 'Box 2', 'Box 3']);
    expect(itemCells.map((tc) => attr(find(tc, 'w:shd'), 'w:fill'))).toEqual(['EF4444', '10B981', 'F59E0B']);
  });

  it('honors the CSS gap as a real blank spacer column between cells, not touching edge to edge', async () => {
    const html = `<div style="display:flex;gap:20px">
      <div style="background:#ef4444">Box 1</div>
      <div style="background:#10b981">Box 2</div>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    // gap:20px -> 15pt -> 300 twips (20 twips/pt). Grid is item, spacer, item.
    const gridCols = findAll(doc, 'w:gridCol');
    expect(gridCols).toHaveLength(3);
    expect(Number(attr(gridCols[1], 'w:w'))).toBe(300);
    const cells = findAll(find(doc, 'w:tr'), 'w:tc');
    expect(cells).toHaveLength(3);
    expect(find(cells[1], 'w:t')).toBeUndefined(); // blank spacer cell, no leaked content/fill
    expect(find(cells[1], 'w:shd')).toBeUndefined();
  });

  it('a multi-row grid gets a real blank spacer row (not just spacer columns) between rows', async () => {
    const html = `<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
      <div style="border:1px solid #ddd">Item 1</div>
      <div style="border:1px solid #ddd">Item 2</div>
      <div style="border:1px solid #ddd">Item 3</div>
      <div style="border:1px solid #ddd">Item 4</div>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const rows = findAll(xml['word/document.xml'], 'w:tr');
    // real row, spacer row, real row = 3.
    expect(rows).toHaveLength(3);
    expect(Number(attr(find(find(rows[1], 'w:trPr'), 'w:trHeight'), 'w:val'))).toBe(150); // 10px -> 7.5pt -> 150 twips
    expect(find(rows[1], 'w:t')).toBeUndefined();
  });

  it('a display:grid container wraps cells into grid-template-columns\' own column count, padding a short final row', async () => {
    const html = `<div style="display:grid;grid-template-columns:1fr 1fr">
      <div style="border:1px solid #ddd">Item 1</div>
      <div style="border:1px solid #ddd">Item 2</div>
      <div style="border:1px solid #ddd">Item 3</div>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const rows = findAll(doc, 'w:tr');
    // real row, spacer row, real row = 3.
    expect(rows).toHaveLength(3);
    // item, spacer, item per real row.
    expect(findAll(rows[0], 'w:tc')).toHaveLength(3);
    const secondRowCells = findAll(rows[2], 'w:tc');
    expect(secondRowCells).toHaveLength(3);
    expect(textContent(find(secondRowCells[0], 'w:t'))).toBe('Item 3');
    expect(find(secondRowCells[2], 'w:t')).toBeUndefined(); // padded empty cell, no border/shd leaked onto it
    const firstCellBorders = find(find(secondRowCells[0], 'w:tcPr'), 'w:tcBorders');
    expect(attr(find(firstCellBorders, 'w:top'), 'w:color')).toBe('DDDDDD');
  });

  it('a multi-block card nested inside a flex/grid cell keeps all its own blocks together in that one cell', async () => {
    const html = `<div style="display:flex;gap:20px">
      <div style="border:1px solid #ddd;padding:10px"><h3>Title</h3><p>Body text.</p></div>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const cells = findAll(doc, 'w:tc');
    expect(cells).toHaveLength(1);
    expect(findAll(cells[0], 'w:p').map((p) => textContent(find(p, 'w:t')))).toEqual(['Title', 'Body text.']);
  });
});

describe('DOCX exporter — rounded/shadowed boxes as real DrawingML shapes', () => {
  function shapeTexts(wsp) {
    return findAll(wsp, 'w:t').map(textContent);
  }

  it('a bordered, rounded, shadowed card renders as a wps:wsp shape, not a plain paragraph border', async () => {
    const html = `<div style="border:1px solid #ddd;padding:20px;border-radius:10px;box-shadow:0 3px 10px rgba(0,0,0,.1)">
      <h3>User Card</h3>
      <p>Name: John Doe</p>
      <p>Role: Software Engineer</p>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shape = find(doc, 'wps:wsp');
    expect(shape).toBeTruthy();
    expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('roundRect');
    expect(attr(find(find(shape, 'a:ln'), 'a:srgbClr'), 'val')).toBe('DDDDDD');
    expect(find(shape, 'a:effectLst')).toBeTruthy();
    expect(shapeTexts(shape)).toEqual(['User Card', 'Name: John Doe', 'Role: Software Engineer']);
    // Exactly one shape holds all three blocks, not three separate ones.
    expect(findAll(doc, 'wps:wsp')).toHaveLength(1);
  });

  it('border-radius:50% on a bordered box renders as a real ellipse, not a rounded rectangle', async () => {
    const html = '<div style="border:1px solid #ddd;border-radius:50%;padding:20px">Circle</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shape = find(doc, 'wps:wsp');
    expect(shape).toBeTruthy();
    expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('ellipse');
  });

  it('a plain border-radius under 50% still renders as roundRect, not ellipse (no regression)', async () => {
    const html = '<div style="border:1px solid #ddd;border-radius:10px;padding:20px">Rounded</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shape = find(doc, 'wps:wsp');
    expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('roundRect');
  });

  it('a borderless rounded gradient panel renders a real a:gradFill, not the flat first-color fallback', async () => {
    const html = '<div style="padding:30px;border-radius:12px;color:white;background:linear-gradient(90deg,#3B82F6,#9333EA)">Gradient Background Example</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shape = find(doc, 'wps:wsp');
    expect(shape).toBeTruthy();
    expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('roundRect');
    const gradFill = find(shape, 'a:gradFill');
    expect(gradFill).toBeTruthy();
    const stops = findAll(gradFill, 'a:gs');
    expect(stops.map((gs) => attr(find(gs, 'a:srgbClr'), 'val'))).toEqual(['3B82F6', '9333EA']);
    expect(shapeTexts(shape)).toEqual(['Gradient Background Example']);
  });

  it('a plain flat-bordered box with no radius/shadow still renders as an ordinary paragraph border (no regression)', async () => {
    const html = '<div style="border:1px solid #ddd;padding:10px"><p>Plain box</p></div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    expect(find(doc, 'wps:wsp')).toBeFalsy();
    expect(find(doc, 'w:pBdr')).toBeTruthy();
  });

  it('a gradient with no border-radius/box-shadow still falls back to a flat fill (no regression)', async () => {
    const html = '<div style="background:linear-gradient(90deg,#ffd,#cdf)">Gradient</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    expect(find(doc, 'wps:wsp')).toBeFalsy();
    expect(attr(find(find(doc, 'w:pPr'), 'w:shd'), 'w:fill')).toBe('FFFFDD');
  });

  it('a page-level wrapper\'s own shadow/border-radius does not turn every plain heading inside it into its own shadowed shape', async () => {
    // Regression: `.slide{background:white;border-radius:15px;box-shadow:...}`
    // (no border of its own) bleeds backgroundColor/rounded/shadow onto every
    // plain, otherwise-unstyled child — only a genuinely self-bordered or
    // self-gradient child should ever be promoted to its own shape.
    const html = `<div style="background:white;border-radius:15px;box-shadow:0 5px 15px rgba(0,0,0,.15);padding:20px">
      <h2>Section</h2>
      <h3>Short Heading</h3>
      <p>Some body text.</p>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    expect(find(doc, 'wps:wsp')).toBeFalsy();
    expect(findAll(doc, 'w:t').map(textContent)).toEqual(expect.arrayContaining(['Section', 'Short Heading', 'Some body text.']));
  });

  it('a shrink-to-fit badge with its own border-radius renders as a real roundRect shape, not a square-cornered table', async () => {
    // A plain auto-fit table (renderShrinkToFitBox) has no OOXML concept of
    // rounded table-cell corners at all, so a pill-shaped badge like this
    // one (border-radius comfortably more than half the badge's own height)
    // rendered with square corners — visibly different from both the source
    // CSS and PPTX's own roundRect shape for the exact same badge.
    const html = '<p style="display:inline-block;padding:6px 12px;background:#2563EB;color:white;border-radius:20px">Badge</p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shape = find(doc, 'wps:wsp');
    expect(shape).toBeTruthy();
    expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('roundRect');
    expect(find(doc, 'w:tbl')).toBeFalsy();
  });

  it('a shrink-to-fit badge with no rounding at all still uses the plain auto-fit table cell', async () => {
    const html = '<p style="display:inline-block;padding:6px 12px;background:#2563EB;color:white">Badge</p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    expect(find(doc, 'wps:wsp')).toBeFalsy();
    expect(find(doc, 'w:tbl')).toBeTruthy();
  });

  it('a shrink-to-fit circular badge renders as a real ellipse shape, not an auto-fit table', async () => {
    const html = '<div style="width:120px;height:120px;border-radius:50%;background:#2563EB;color:white">72%</div>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shape = find(doc, 'wps:wsp');
    expect(shape).toBeTruthy();
    expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('ellipse');
    expect(find(doc, 'w:tbl')).toBeFalsy();
    expect(findAll(shape, 'w:t').map(textContent)).toEqual(['72%']);
    // Explicit width/height is honored directly, not the font-metrics guess.
    const ext = find(shape, 'a:ext');
    expect(Number(attr(ext, 'cx'))).toBe(120 * 9525);
    expect(Number(attr(ext, 'cy'))).toBe(120 * 9525);
    // Regression: wrap:'none' is NOT used for a Word wps:wsp shape (unlike
    // PPTX's native p:sp) — observed breaking text into one character per
    // line in real-world renderers. Must be 'square', and the label must be
    // centered since the source declared no explicit text-align.
    const bodyPr = find(shape, 'wps:bodyPr');
    expect(attr(bodyPr, 'wrap')).toBe('square');
    const jc = find(find(find(shape, 'w:p'), 'w:pPr'), 'w:jc');
    expect(attr(jc, 'w:val')).toBe('center');
  });

  it('an unsized circular badge falls back to the font-metrics single-line estimate', async () => {
    const html = '<p style="display:inline-block;border-radius:50%;background:#2563EB;color:white">7</p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shape = find(doc, 'wps:wsp');
    expect(shape).toBeTruthy();
    expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('ellipse');
    const ext = find(shape, 'a:ext');
    expect(Number(attr(ext, 'cx'))).toBeGreaterThan(0);
    expect(Number(attr(ext, 'cy'))).toBeGreaterThan(0);
  });

  it('the card group\'s own outer margin spaces the shape as a whole, not each internal paragraph individually', async () => {
    const html = `<div style="border:1px solid #ddd;border-radius:10px;margin:20px 0">
      <h3>Title</h3>
      <p>Body</p>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shape = find(doc, 'wps:wsp');
    // 20px -> 15pt -> 300 twips, applied once to the shape's own w:p.
    const drawingParagraph = findAll(doc, 'w:p').find((p) => find(p, 'w:drawing'));
    expect(attr(find(find(drawingParagraph, 'w:pPr'), 'w:spacing'), 'w:before')).toBe('300');
    // Not duplicated onto the members inside the shape's own text body.
    const innerParagraphs = findAll(shape, 'w:p');
    for (const p of innerParagraphs) {
      const pPr = find(p, 'w:pPr');
      expect(pPr ? find(pPr, 'w:spacing') : undefined).toBeFalsy();
    }
  });

  it('a flex item with its own border-radius renders as a shape sized to its own column, not a flat table cell', async () => {
    const html = `<div style="display:flex;gap:20px">
      <div style="flex:1;padding:20px;border-radius:8px;color:white;background:#ef4444">Box 1</div>
      <div style="flex:1;padding:20px;border-radius:8px;color:white;background:#10b981">Box 2</div>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shapes = findAll(doc, 'wps:wsp');
    expect(shapes).toHaveLength(2);
    for (const shape of shapes) {
      expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('roundRect');
    }
    expect(shapeTexts(shapes[0])).toEqual(['Box 1']);
    expect(shapeTexts(shapes[1])).toEqual(['Box 2']);
    // Row layout (side-by-side positioning via the table) is unaffected.
    expect(findAll(doc, 'w:tr')).toHaveLength(1);
  });

  it('a grid item with border-radius and box-shadow renders as a shadowed shape inside its cell', async () => {
    const html = `<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
      <div style="border:1px solid #ddd;padding:20px;border-radius:10px;box-shadow:0 3px 10px rgba(0,0,0,.1)">Grid Item 1</div>
      <div style="border:1px solid #ddd;padding:20px;border-radius:10px;box-shadow:0 3px 10px rgba(0,0,0,.1)">Grid Item 2</div>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    const shapes = findAll(doc, 'wps:wsp');
    expect(shapes).toHaveLength(2);
    for (const shape of shapes) {
      expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('roundRect');
      expect(find(shape, 'a:effectLst')).toBeTruthy();
    }
  });

  it('a flex/grid item with a plain flat border (no radius/shadow) still renders as an ordinary table cell (no regression)', async () => {
    const html = `<div style="display:flex;gap:20px">
      <div style="border:1px solid #ddd">Item 1</div>
      <div style="border:1px solid #ddd">Item 2</div>
    </div>`;
    const { buffer } = await convert(html, { format: 'docx' });
    const { xml } = readDocxParts(buffer);
    const doc = xml['word/document.xml'];
    expect(find(doc, 'wps:wsp')).toBeFalsy();
    expect(findAll(doc, 'w:tcBorders').length).toBeGreaterThan(0);
  });
});
