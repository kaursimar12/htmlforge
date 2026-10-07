import { describe, it, expect } from 'vitest';
import { deflateSync } from 'node:zlib';
import { convert } from '../src/index.js';
import {
  parsePdf, getCatalog, getPages, getPageContentOps, getAnnotations, resolveRef, resolveResource,
} from './helpers/pdf.js';

// The bullet glyph is drawn under WinAnsiEncoding (Windows-1252), whose
// bullet code point is byte 0x95 — the raw Unicode '•' is U+2022 and would
// never appear verbatim in a decoded content stream (see content-stream.js's
// encodeWinAnsi).
const BULLET_WINANSI = String.fromCharCode(0x95);

// ─── Synthetic image fixtures ───────────────────────────────────────────
// Real PNG/JPEG bytes small enough to inline, built by hand (this project
// has no image library) — just enough structure for this exporter's own
// parsers (images.js) to extract what they need; not necessarily a
// pixel-perfect renderable image in a real viewer.

function crc32(buf) {
  if (!crc32.table) {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    crc32.table = table;
  }
  let c = 0xffffffff;
  for (const byte of buf) c = crc32.table[(c ^ byte) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const typeBuf = Buffer.from(type, 'ascii');
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.length, 0);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
}

function buildTestPng({ width = 2, height = 2, colorType = 2 } = {}) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = colorType;
  const channels = colorType === 2 ? 3 : colorType === 0 ? 1 : 4; // enough for the color types this fixture builder is asked for
  const rowBytes = 1 + width * channels;
  const raw = Buffer.alloc(rowBytes * height, 128);
  for (let y = 0; y < height; y++) raw[y * rowBytes] = 0; // filter type 0 (None) per row
  const idat = deflateSync(raw);
  return Buffer.concat([sig, pngChunk('IHDR', ihdr), pngChunk('IDAT', idat), pngChunk('IEND', Buffer.alloc(0))]);
}

// A minimal colorType-3 (palette) PNG — every pixel is index 0, PLTE's own
// single entry is a mid-gray RGB triplet, matching buildTestPng's own
// "solid fill, filter type None" simplicity.
function buildTestIndexedPng({ width = 2, height = 2 } = {}) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 3; // palette
  const rowBytes = 1 + width; // 1 index byte per pixel
  const raw = Buffer.alloc(rowBytes * height, 0); // every pixel = palette index 0
  for (let y = 0; y < height; y++) raw[y * rowBytes] = 0; // filter type 0 (None)
  const idat = deflateSync(raw);
  const plte = Buffer.from([128, 128, 128]); // one palette entry: mid-gray
  return Buffer.concat([sig, pngChunk('IHDR', ihdr), pngChunk('PLTE', plte), pngChunk('IDAT', idat), pngChunk('IEND', Buffer.alloc(0))]);
}

function buildTestJpeg({ width = 10, height = 8, components = 3 } = {}) {
  const soi = Buffer.from([0xff, 0xd8]);
  const sofData = Buffer.alloc(6 + components * 3);
  let o = 0;
  sofData[o++] = 8; // precision
  sofData.writeUInt16BE(height, o); o += 2;
  sofData.writeUInt16BE(width, o); o += 2;
  sofData[o++] = components;
  for (let i = 0; i < components; i++) { sofData[o++] = i + 1; sofData[o++] = 0x11; sofData[o++] = 0; }
  const lengthField = Buffer.alloc(2);
  lengthField.writeUInt16BE(sofData.length + 2, 0);
  const sof0 = Buffer.concat([Buffer.from([0xff, 0xc0]), lengthField, sofData]);
  const eoi = Buffer.from([0xff, 0xd9]);
  return Buffer.concat([soi, sof0, eoi]);
}

function dataUri(mimeType, buffer) {
  return `data:${mimeType};base64,${buffer.toString('base64')}`;
}

async function convertPdf(html, options) {
  return convert(html, { ...options, format: 'pdf' });
}

describe('PDF exporter', () => {
  it('returns correct mimeType and extension', async () => {
    const result = await convertPdf('<p>hi</p>');
    expect(result.mimeType).toBe('application/pdf');
    expect(result.extension).toBe('.pdf');
  });

  describe('PDF1 — skeleton', () => {
    it('produces a valid PDF header, xref table, and trailer', async () => {
      const { buffer } = await convertPdf('<p>hi</p>');
      expect(buffer.subarray(0, 5).toString('latin1')).toBe('%PDF-');
      const pdf = parsePdf(buffer);
      expect(pdf.trailer.Root).toBeTruthy();
    });

    it('trailer Root resolves to a Catalog whose Pages resolves with a correct Count', async () => {
      const { buffer } = await convertPdf('<p>one</p><p>two</p>');
      const pdf = parsePdf(buffer);
      const catalog = getCatalog(pdf);
      expect(catalog.Type).toBe('Catalog');
      const pagesRoot = resolveRef(pdf, catalog.Pages);
      expect(pagesRoot.Type).toBe('Pages');
      expect(pagesRoot.Count).toBe(pagesRoot.Kids.length);
    });

    it('an empty document still produces exactly one page with a correct MediaBox', async () => {
      const { buffer } = await convertPdf('');
      const pdf = parsePdf(buffer);
      const pages = getPages(pdf);
      expect(pages).toHaveLength(1);
      expect(pages[0].MediaBox).toEqual([0, 0, 595, 842]);
    });
  });

  describe('PDF2 — text / paragraph flow', () => {
    it('a paragraph draws via BT/Tf/Tj with a resolvable font resource and round-tripped text', async () => {
      const { buffer } = await convertPdf('<p>Hello there</p>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      expect(ops.some((o) => o.op === 'BT')).toBe(true);
      expect(ops.some((o) => o.op === 'ET')).toBe(true);
      const tf = ops.find((o) => o.op === 'Tf');
      expect(tf).toBeTruthy();
      const font = resolveResource(pdf, page, 'Font', tf.operands[0]);
      expect(font.Subtype).toBe('Type1');
      const text = ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0]).join(' ');
      expect(text).toContain('Hello');
    });

    it('a long paragraph wraps across multiple lines', async () => {
      const longText = Array.from({ length: 40 }, (_, i) => `word${i}`).join(' ');
      const { buffer } = await convertPdf(`<p>${longText}</p>`);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const tmYs = new Set(ops.filter((o) => o.op === 'Tm').map((o) => o.operands[5]));
      expect(tmYs.size).toBeGreaterThan(1);
    });

    it('a <div style="white-space:pre"> code-lookalike (not a real <pre>) draws one shared background inset by padding, preserving a blank line between statements as vertical space', async () => {
      const html = '<div style="background:#1e293b;color:#fff;white-space:pre">\nconst pdf = generate(html);\n\nconsole.log(pdf);\n</div>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const texts = ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      // Two separate Tj calls, not one run-on line joining both statements.
      expect(texts).toEqual(['const pdf = generate(html);', 'console.log(pdf);']);
      // One shared background rect spans the whole block (top/bottom inset by
      // its own padding), not a separate bar per source line — including the
      // blank line, which still occupies its own line-height of vertical
      // space between the two statements rather than collapsing away.
      expect(ops.filter((o) => o.op === 'f')).toHaveLength(1);
      // The two statements sit two line-heights apart (their own line, plus
      // the blank line's), not one — confirming the blank line still
      // occupies real vertical space rather than collapsing away now that
      // there's no longer a separate empty bar marking its presence.
      const tmYs = ops.filter((o) => o.op === 'Tm').map((o) => o.operands[5]);
      const defaultLineHeightPt = 11 * 1.2; // this test's HTML declares no body font-size/line-height
      expect(tmYs[0] - tmYs[1]).toBeCloseTo(defaultLineHeightPt * 2, 1);
    });
  });

  describe('PDF3 — headings', () => {
    it.each([1, 2, 3, 4, 5, 6])('heading level %i uses its documented font-size tier and a bold font', async (level) => {
      const { buffer } = await convertPdf(`<h${level}>Heading</h${level}>`);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const tf = ops.find((o) => o.op === 'Tf');
      const expectedSize = { 1: 24, 2: 18, 3: 14, 4: 12, 5: 11, 6: 10 }[level];
      expect(tf.operands[1]).toBe(expectedSize);
      const font = resolveResource(pdf, page, 'Font', tf.operands[0]);
      expect(font.BaseFont).toContain('Bold');
    });
  });

  describe('PDF4 — inline formatting', () => {
    it('mixed plain/bold/italic/code text switches fonts at the right word boundaries', async () => {
      const { buffer } = await convertPdf('<p>plain <b>bold</b> <i>italic</i> <code>code</code></p>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);

      const fontNamesInOrder = ops.filter((o) => o.op === 'Tf').map((o) => o.operands[0]);
      const baseFontsInOrder = fontNamesInOrder.map((rn) => resolveResource(pdf, page, 'Font', rn).BaseFont);
      expect(baseFontsInOrder).toEqual(['Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique', 'Courier']);

      const textInOrder = ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(textInOrder).toEqual(['plain', 'bold', 'italic', 'code']);
    });

    // Regression test: flattenRuns's styledSpan accumulator (blocks.js)
    // previously dropped `fontSizePt` while forwarding every other inherited
    // style field (bold/italic/color/fontFamily/underline/...), so an
    // explicit CSS font-size anywhere below the top-level block (an inline
    // <span>, or a <p>/<td>'s own style — both wrap their content in the
    // same styledSpan, see dom-to-idm.js's mapInlineWithStyle) silently
    // rendered at the surrounding block's own base size instead.
    it('an inline <span> with its own font-size draws at that size, not the paragraph\'s base size', async () => {
      const { buffer } = await convertPdf('<p>normal <span style="font-size:24px">big</span> normal again</p>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const tfSizes = ops.filter((o) => o.op === 'Tf').map((o) => o.operands[1]);
      expect(tfSizes).toEqual([11, 18, 11]); // 24px * 0.75 = 18pt; default body size 11pt either side
    });

    it('a <p> with its own font-size draws its whole text at that size', async () => {
      const { buffer } = await convertPdf('<p style="font-size:20px">Big paragraph</p><p>Normal paragraph</p>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const tfSizes = ops.filter((o) => o.op === 'Tf').map((o) => o.operands[1]);
      expect(tfSizes).toEqual([15, 11]); // 20px * 0.75 = 15pt, then the second paragraph's default 11pt
    });

    it('a table cell with its own font-size draws its text at that size, independent of other cells', async () => {
      const { buffer } = await convertPdf('<table><tr><td style="font-size:24px">Big cell</td><td>Normal cell</td></tr></table>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const tfSizes = ops.filter((o) => o.op === 'Tf').map((o) => o.operands[1]);
      expect(tfSizes).toEqual(expect.arrayContaining([18, 11]));
    });

    // Regression test: width ESTIMATION (used to position the next segment)
    // never accounted for a <code> run always drawing in Courier — a wider,
    // fixed-width font than whatever ambient/proportional font surrounds it
    // — so the segment right after a <code> span was positioned using an
    // underestimated width and visibly overlapped the code span's own tail.
    it('text after an inline <code> span is positioned correctly, not overlapping it', async () => {
      const { buffer } = await convertPdf('<p>before <code>&lt;link&gt;</code> after</p>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const positions = ops.filter((o) => o.op === 'Tm').map((o) => o.operands[4]);
      const texts = ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(texts).toEqual(['before', '<link>', 'after']);

      // The real drawn width of "<link>" is 6 Courier characters at the
      // default 11pt body size: 6 * 0.6 * 11 = 39.6pt — "after" must start
      // at least that far past where "<link>" itself starts, not overlap it.
      const codeStartX = positions[1];
      const afterStartX = positions[2];
      expect(afterStartX - codeStartX).toBeGreaterThanOrEqual(39.6 - 0.01);
    });
  });

  describe('PDF5 — lists', () => {
    it('an unordered list draws a bullet marker for each item', async () => {
      const { buffer } = await convertPdf('<ul><li>Apple</li><li>Orange</li></ul>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const texts = ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      // The bullet glyph is drawn via WinAnsiEncoding (Windows-1252), whose
      // bullet code point is byte 0x95 — not the raw Unicode '•' (U+2022).
      expect(texts.filter((t) => t === BULLET_WINANSI)).toHaveLength(2);
      expect(texts).toContain('Apple');
      expect(texts).toContain('Orange');
    });

    it('an ordered list with start=5 numbers items starting at 5', async () => {
      const { buffer } = await convertPdf('<ol start="5"><li>First</li><li>Second</li></ol>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const texts = ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(texts).toContain('5.');
      expect(texts).toContain('6.');
    });

    it('a nested list indents its own items further than the parent', async () => {
      const { buffer } = await convertPdf('<ul><li>Top<ul><li>Nested</li></ul></li></ul>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);

      // Each bullet's "•" Tj is immediately preceded by its own positioning
      // Tm — pairing them up gives the x-position each bullet was actually
      // drawn at, in document order.
      const bulletTmXs = [];
      for (let k = 0; k < ops.length - 1; k++) {
        if (ops[k].op === 'Tm' && ops[k + 1].op === 'Tj' && ops[k + 1].operands[0] === BULLET_WINANSI) bulletTmXs.push(ops[k].operands[4]);
      }
      expect(bulletTmXs).toHaveLength(2);
      expect(bulletTmXs[1]).toBeGreaterThan(bulletTmXs[0]);
    });
  });

  describe('PDF6 — page breaks', () => {
    it('enough content to overflow one page produces multiple pages with no dropped content', async () => {
      const paragraphs = Array.from({ length: 80 }, (_, i) => `<p>Paragraph number ${i} with some real content in it.</p>`).join('');
      const { buffer } = await convertPdf(paragraphs);
      const pdf = parsePdf(buffer);
      const pages = getPages(pdf);
      expect(pages.length).toBeGreaterThan(1);
      const allText = pages.flatMap((p) => getPageContentOps(pdf, p).filter((o) => o.op === 'Tj').map((o) => o.operands[0])).join(' ');
      expect(allText).toContain('Paragraph number 0');
      expect(allText).toContain('Paragraph number 79');
    });

    it('an explicit page break forces a new page even with room left on the current one', async () => {
      const { buffer } = await convertPdf('<p>Before</p><hr class="page-break"><p>After</p>');
      const pdf = parsePdf(buffer);
      const pages = getPages(pdf);
      expect(pages).toHaveLength(2);
      const firstPageText = getPageContentOps(pdf, pages[0]).filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      const secondPageText = getPageContentOps(pdf, pages[1]).filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(firstPageText).toContain('Before');
      expect(secondPageText).toContain('After');
    });
  });

  describe('PDF7 — images', () => {
    it('a JPEG embeds via DCTDecode with the real byte length', async () => {
      const jpeg = buildTestJpeg({ width: 10, height: 8 });
      const { buffer } = await convertPdf(`<img src="${dataUri('image/jpeg', jpeg)}" alt="x">`);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const doOp = ops.find((o) => o.op === 'Do');
      expect(doOp).toBeTruthy();
      const xobj = resolveResource(pdf, page, 'XObject', doOp.operands[0]);
      expect(xobj.Filter).toBe('DCTDecode');
      expect(xobj.Width).toBe(10);
      expect(xobj.Height).toBe(8);
      expect(xobj.__stream.length).toBe(jpeg.length);
    });

    it('a non-interlaced 8-bit RGB PNG (no alpha) decodes and embeds via plain FlateDecode, no SMask', async () => {
      const png = buildTestPng({ width: 2, height: 2, colorType: 2 });
      const { buffer, warnings } = await convertPdf(`<img src="${dataUri('image/png', png)}" alt="x">`);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const doOp = ops.find((o) => o.op === 'Do');
      expect(doOp).toBeTruthy();
      const xobj = resolveResource(pdf, page, 'XObject', doOp.operands[0]);
      expect(xobj.Filter).toBe('FlateDecode');
      expect(xobj.ColorSpace).toBe('DeviceRGB');
      expect(xobj.SMask).toBeUndefined();
      expect(warnings.filter((w) => w.code === 'pdf-unsupported-png')).toHaveLength(0);
    });

    it('an RGBA PNG (colorType 6, has alpha) decodes, embeds, and carries its alpha as a real /SMask', async () => {
      const png = buildTestPng({ width: 2, height: 2, colorType: 6 });
      const { buffer, warnings } = await convertPdf(`<img src="${dataUri('image/png', png)}" alt="x">`);
      expect(warnings.filter((w) => w.code === 'pdf-unsupported-png')).toHaveLength(0);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const doOp = ops.find((o) => o.op === 'Do');
      expect(doOp).toBeTruthy();
      const xobj = resolveResource(pdf, page, 'XObject', doOp.operands[0]);
      expect(xobj.ColorSpace).toBe('DeviceRGB');
      expect(xobj.SMask).toBeTruthy();
      const smask = resolveRef(pdf, xobj.SMask);
      expect(smask.ColorSpace).toBe('DeviceGray');
      expect(smask.Width).toBe(2);
      expect(smask.Height).toBe(2);
    });

    it('a palette (colorType 3) PNG expands via PLTE to real RGB', async () => {
      const png = buildTestIndexedPng({ width: 2, height: 2 });
      const { buffer, warnings } = await convertPdf(`<img src="${dataUri('image/png', png)}" alt="x">`);
      expect(warnings.filter((w) => w.code === 'pdf-unsupported-png')).toHaveLength(0);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const doOp = ops.find((o) => o.op === 'Do');
      expect(doOp).toBeTruthy();
      const xobj = resolveResource(pdf, page, 'XObject', doOp.operands[0]);
      expect(xobj.ColorSpace).toBe('DeviceRGB');
    });

    it('a declared width/height attribute is ignored in favor of the image\'s own real intrinsic size', async () => {
      // A stale `<img width="500" height="500">` pointing at a real 2x2
      // image must render at the image's own actual size, not the claimed
      // one — see images.js/blocks.js's own "real beats declared" precedence.
      const png = buildTestPng({ width: 2, height: 2, colorType: 2 });
      const { buffer } = await convertPdf(`<img src="${dataUri('image/png', png)}" width="500" height="500" alt="x">`);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const doIndex = ops.findIndex((o) => o.op === 'Do');
      const cmOp = [...ops.slice(0, doIndex)].reverse().find((o) => o.op === 'cm');
      // 2px * 72/96 PX_TO_PT = 1.5pt drawn width/height, not 500px's 375pt.
      expect(cmOp.operands[0]).toBeCloseTo(1.5, 5);
      expect(cmOp.operands[3]).toBeCloseTo(1.5, 5);
    });

    it('a truncated/corrupt PNG is skipped with a warning instead of crashing the export', async () => {
      const png = buildTestPng({ width: 4, height: 4, colorType: 2 });
      const truncated = png.subarray(0, png.length - 10);
      const { buffer, warnings } = await convertPdf(`<p>Before</p><img src="${dataUri('image/png', truncated)}" alt="x"><p>After</p>`);
      expect(buffer.subarray(0, 5).toString('latin1')).toBe('%PDF-');
      expect(warnings.some((w) => w.code === 'pdf-corrupt-image')).toBe(true);
    });

    it('an <img> alone in a table cell (no surrounding text) still draws as a real XObject', async () => {
      const png = buildTestPng({ width: 2, height: 2, colorType: 2 });
      const html = `<table><tr><td>Label</td><td><img src="${dataUri('image/png', png)}" alt="x"></td></tr></table>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const doOp = ops.find((o) => o.op === 'Do');
      expect(doOp).toBeTruthy();
      const xobj = resolveResource(pdf, page, 'XObject', doOp.operands[0]);
      expect(xobj.Width).toBe(2);
      expect(xobj.Height).toBe(2);
    });

    it('a table cell mixing text and an <img> draws both the text and the image', async () => {
      const png = buildTestPng({ width: 2, height: 2, colorType: 2 });
      const html = `<table><tr><td><b>Caption</b><img src="${dataUri('image/png', png)}" alt="x"></td></tr></table>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      expect(ops.some((o) => o.op === 'Do')).toBe(true);
      expect(ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0])).toContain('Caption');
    });

    it('a cell holding only an image (no text) centers it vertically within its own row when the row is taller than the image needs', async () => {
      // A short, wide image next to a sibling cell with several lines of
      // text — the ROW's own height is driven by the TEXT cell, leaving
      // slack above/below the image cell's own (shorter) content.
      const png = buildTestPng({ width: 100, height: 20, colorType: 2 });
      const html = `<table><tr><td>Line one<br>Line two<br>Line three<br>Line four</td><td><img src="${dataUri('image/png', png)}" alt="x"></td></tr></table>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);

      const rowRect = ops.find((o) => o.op === 're');
      const rowBottom = rowRect.operands[1];
      const rowTop = rowBottom + rowRect.operands[3];

      const doIndex = ops.findIndex((o) => o.op === 'Do');
      const cmOp = [...ops.slice(0, doIndex)].reverse().find((o) => o.op === 'cm');
      const imgBottom = cmOp.operands[5];
      const imgTop = imgBottom + cmOp.operands[3];

      const topGap = rowTop - imgTop;
      const bottomGap = imgBottom - rowBottom;
      expect(Math.abs(topGap - bottomGap)).toBeLessThan(1);
    });
  });

  describe('SVG vector shapes', () => {
    it('a rect+circle logo (no viewBox) draws a filled rect then a filled ellipse on top, in the right relative position', async () => {
      const html = '<svg width="200" height="120"><rect width="200" height="120" fill="#2563eb"/><circle cx="100" cy="60" r="35" fill="orange"/></svg>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);

      const reIndex = ops.findIndex((o) => o.op === 're');
      expect(reIndex).toBeGreaterThanOrEqual(0);
      const rectFill = [...ops.slice(reIndex)].find((o) => o.op === 'rg')?.operands;
      expect(rectFill[0]).toBeCloseTo(0x25 / 255, 2);

      // The circle draws as a real Bézier ellipse path (no native PDF
      // ellipse/circle operator exists), after the rect.
      const curveIndex = ops.findIndex((o) => o.op === 'c');
      expect(curveIndex).toBeGreaterThan(reIndex);
      const circleFill = [...ops.slice(curveIndex)].find((o) => o.op === 'rg')?.operands;
      // orange = #FFA500, rounded to 3 decimals by content-stream.js's own number formatting.
      expect(circleFill[0]).toBeCloseTo(1, 2);
      expect(circleFill[1]).toBeCloseTo(0xa5 / 0xff, 2);
      expect(circleFill[2]).toBeCloseTo(0, 2);

      // The circle (bounding box centered at 100,60 in a 200x120 box) sits
      // horizontally centered and slightly below vertical center relative
      // to the rect — never at the rect's own corner/edge.
      const rectBounds = ops[reIndex].operands;
      const curveOps = ops.filter((o) => o.op === 'c' || o.op === 'm');
      const circleXs = curveOps.map((o) => o.operands[0]);
      const circleCenterX = (Math.min(...circleXs) + Math.max(...circleXs)) / 2;
      expect(circleCenterX).toBeCloseTo(rectBounds[0] + rectBounds[2] / 2, 0);
    });

    it('a viewBox-scaled icon (e.g. a 24x24 icon shown at 48x48) rescales shape geometry accordingly', async () => {
      const html = '<svg width="48" height="48" viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="20" fill="red"/></svg>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const rect = ops.find((o) => o.op === 're');
      // 20 raw units * (48/24 viewBox scale) * (72/96 px-to-pt) = 30pt.
      expect(rect.operands[2]).toBeCloseTo(30, 1);
      expect(rect.operands[3]).toBeCloseTo(30, 1);
    });

    it('a line draws a stroke with no fill; a polygon closes for both fill and stroke', async () => {
      const html = '<svg width="100" height="100"><line x1="0" y1="0" x2="100" y2="100" stroke="black"/><polygon points="0,0 10,0 5,10" fill="red" stroke="blue"/></svg>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      expect(ops.some((o) => o.op === 'S')).toBe(true); // the line's stroke-only paint
      expect(ops.some((o) => o.op === 'B')).toBe(true); // the polygon's fill+stroke paint
      expect(ops.some((o) => o.op === 'h')).toBe(true); // the polygon's explicit close
    });

    it('a shape with fill="none" and no stroke draws nothing at all', async () => {
      const html = '<svg width="50" height="50"><rect width="50" height="50" fill="none"/></svg>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      expect(ops.some((o) => o.op === 're' || o.op === 'f' || o.op === 'S' || o.op === 'B')).toBe(false);
    });

    it('an SVG with only unsupported content (e.g. a <path>) still drops silently, same as before', async () => {
      const { buffer } = await convertPdf('<svg><path d="M10 10 L90 90"/></svg><p>after</p>');
      const pdf = parsePdf(buffer);
      const texts = getPageContentOps(pdf, getPages(pdf)[0]).filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(texts).toEqual(['after']);
    });
  });

  describe('PDF8 — tables', () => {
    it('a table draws a filled/bordered rect per cell with bold header text', async () => {
      const html = '<table><tr><th>Name</th><th>Age</th></tr><tr><td>Alice</td><td>30</td></tr></table>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const ops = getPageContentOps(pdf, page);
      const reOps = ops.filter((o) => o.op === 're');
      expect(reOps.length).toBeGreaterThanOrEqual(4); // 2 cols x 2 rows
      const texts = ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(texts).toEqual(expect.arrayContaining(['Name', 'Age', 'Alice', '30']));

      const fontNames = ops.filter((o) => o.op === 'Tf').map((o) => o.operands[0]);
      const baseFonts = fontNames.map((rn) => resolveResource(pdf, page, 'Font', rn).BaseFont);
      expect(baseFonts).toContain('Helvetica-Bold');
    });

    it('a rowspan/colspan table expands to a consistent grid without throwing', async () => {
      const html = `<table>
        <tr><td rowspan="2">A</td><td>B</td></tr>
        <tr><td>C</td></tr>
        <tr><td colspan="2">D</td></tr>
      </table>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const texts = getPageContentOps(pdf, page).filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(texts).toEqual(expect.arrayContaining(['A', 'B', 'C', 'D']));
    });

    it('a table long enough to span a page break repeats its header row on the continuation page', async () => {
      const rows = Array.from({ length: 60 }, (_, i) => `<tr><td>Row ${i}</td><td>Value ${i}</td></tr>`).join('');
      const html = `<table><thead><tr><th>Col A</th><th>Col B</th></tr></thead><tbody>${rows}</tbody></table>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const pages = getPages(pdf);
      expect(pages.length).toBeGreaterThan(1);
      const secondPageTexts = getPageContentOps(pdf, pages[1]).filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(secondPageTexts).toContain('Col A');
    });

    it('a header cell with its own explicit background color uses that color, not the default gray fallback', async () => {
      const html = '<table><tr><th style="background:#2563eb;color:#fff">Quarter</th><th>Revenue</th></tr><tr><td>Q1</td><td>$450K</td></tr></table>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);

      // Each cell draws `q re rg RG w B Q` — the fill color op sits between
      // the path (`re`) and the paint op (`B`), which is valid PDF (paint
      // operators apply whatever color state is current at paint time, not
      // at path-construction time).
      const rectIndices = ops.map((o, i) => (o.op === 're' ? i : -1)).filter((i) => i >= 0);
      const fillFor = (rectIndex) => {
        const paintIndex = ops.findIndex((o, i) => i > rectIndex && (o.op === 'B' || o.op === 'f'));
        return ops.slice(rectIndex, paintIndex).find((o) => o.op === 'rg')?.operands;
      };
      // rounded to 3 decimals by content-stream.js's own number formatting
      const firstFill = fillFor(rectIndices[0]);
      expect(firstFill[0]).toBeCloseTo(0x25 / 255, 2);
      expect(firstFill[1]).toBeCloseTo(0x63 / 255, 2);
      expect(firstFill[2]).toBeCloseTo(0xeb / 255, 2);
      // The sibling header cell (no explicit background) still falls back to
      // the default gray, so the fallback itself isn't broken by this fix.
      expect(fillFor(rectIndices[1])).toEqual([0.9, 0.9, 0.92]);
    });

    it('columns size to their own content, not a flat equal split', async () => {
      // "A very long column header indeed" needs far more room than "X" —
      // a flat equal split would force the long column to wrap or overflow
      // while leaving the short column mostly empty, which a browser's own
      // table auto-layout never does.
      const html = '<table><tr><th>A very long column header indeed</th><th>X</th></tr><tr><td>Some reasonably long cell content</td><td>Y</td></tr></table>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const rects = ops.filter((o) => o.op === 're');
      const [wideCol, narrowCol] = rects; // header row is drawn first, top to bottom, left to right
      expect(wideCol.operands[2]).toBeGreaterThan(narrowCol.operands[2] * 3);
      // The two columns still sum to the full available content width.
      expect(wideCol.operands[2] + narrowCol.operands[2]).toBeCloseTo(451, 0);
    });

    it('an explicit width="%" on every first-row cell distributes by that proportion instead of estimating from content', async () => {
      const html = '<table><tr><td width="75%">Short</td><td width="25%">Also short</td></tr></table>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const rects = ops.filter((o) => o.op === 're');
      const [wideCol, narrowCol] = rects;
      expect(wideCol.operands[2]).toBeCloseTo(451 * 0.75, 0);
      expect(narrowCol.operands[2]).toBeCloseTo(451 * 0.25, 0);
    });

    it('a table whose content overflows the available width scales every column down proportionally rather than overflowing the page', async () => {
      const longText = Array.from({ length: 30 }, (_, i) => `word${i}`).join(' ');
      const html = `<table><tr><td>${longText}</td><td>${longText}</td><td>${longText}</td></tr></table>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const rects = ops.filter((o) => o.op === 're');
      const totalWidth = rects.slice(0, 3).reduce((sum, r) => sum + r.operands[2], 0);
      expect(totalWidth).toBeCloseTo(451, 0);
    });

    // Regression test: the content-based column-width estimator always
    // measured a cell's text against the generic (proportional-font)
    // fallback, completely ignoring the cell's own declared font-family —
    // so a monospace/Courier cell's column was sized using a narrower
    // estimate than its real (flat, wider-per-character) rendered width,
    // and the actual text ran past the column's own right edge.
    it('a column containing a monospace-font cell is sized wide enough for its real content, not the generic estimate', async () => {
      const { estimateTextWidth } = await import('../src/exporter/shared/font-metrics.js');
      const text = 'The quick brown fox jumps over the lazy dog.';
      const html = `<table><tr><td style="font-family:'Courier New'">${text}</td><td>x</td></tr></table>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const [monoCol] = getPageContentOps(pdf, getPages(pdf)[0]).filter((o) => o.op === 're');

      const realWidth = estimateTextWidth(text, 11, 'Courier New');
      expect(monoCol.operands[2]).toBeGreaterThanOrEqual(realWidth);
    });
  });

  describe('PDF9 — hyperlinks', () => {
    it('a hyperlink produces a Link annotation with the correct URI overlapping the drawn text', async () => {
      const { buffer } = await convertPdf('<p>Visit <a href="https://example.com">this site</a> now.</p>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      const annots = getAnnotations(pdf, page);
      expect(annots).toHaveLength(1);
      expect(annots[0].Subtype).toBe('Link');
      expect(annots[0].A.URI).toBe('https://example.com');
      expect(annots[0].Rect).toHaveLength(4);
    });

    it('an internal #anchor link renders as plain text with no annotation', async () => {
      const { buffer } = await convertPdf('<p><a href="#section1">Jump</a></p>');
      const pdf = parsePdf(buffer);
      const [page] = getPages(pdf);
      expect(getAnnotations(pdf, page)).toHaveLength(0);
      const texts = getPageContentOps(pdf, page).filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(texts).toContain('Jump');
    });
  });

  describe('PDF10 — metadata + page setup', () => {
    it('Info reflects title/author overrides and is entirely omitted when absent', async () => {
      const withMeta = await convertPdf('<p>hi</p>', { title: 'My Report', author: 'Jane Doe' });
      const pdfWithMeta = parsePdf(withMeta.buffer);
      expect(pdfWithMeta.trailer.Info).toBeTruthy();
      const info = resolveRef(pdfWithMeta, pdfWithMeta.trailer.Info);
      expect(info.Title).toBe('My Report');
      expect(info.Author).toBe('Jane Doe');

      const withoutMeta = await convertPdf('<p>hi</p>');
      const pdfWithoutMeta = parsePdf(withoutMeta.buffer);
      expect(pdfWithoutMeta.trailer.Info).toBeUndefined();
    });

    it('pageSize/orientation/margins produce the correct MediaBox', async () => {
      const letter = await convertPdf('<p>hi</p>', { pageSize: 'Letter' });
      expect(getPages(parsePdf(letter.buffer))[0].MediaBox).toEqual([0, 0, 612, 792]);

      const a4 = await convertPdf('<p>hi</p>', { pageSize: 'A4' });
      expect(getPages(parsePdf(a4.buffer))[0].MediaBox).toEqual([0, 0, 595, 842]);

      const legalLandscape = await convertPdf('<p>hi</p>', { pageSize: 'Legal', orientation: 'landscape' });
      expect(getPages(parsePdf(legalLandscape.buffer))[0].MediaBox).toEqual([0, 0, 1008, 612]);
    });
  });

  describe('box styling (backgroundColor/border/shadow/rounded/padding/margin/cardGroup/shrinkToFit)', () => {
    // Finds the fill (rg) / stroke (RG) color nearest before a given op index.
    function colorsBefore(ops, index) {
      const fill = [...ops.slice(0, index)].reverse().find((o) => o.op === 'rg');
      const stroke = [...ops.slice(0, index)].reverse().find((o) => o.op === 'RG');
      return { fill: fill?.operands, stroke: stroke?.operands };
    }

    it('a div with a background color draws a filled rect behind its paragraph', async () => {
      const { buffer } = await convertPdf('<div style="background:#ff0000"><p>Hi</p></div>');
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const fillIndex = ops.findIndex((o) => o.op === 'f');
      const tjIndex = ops.findIndex((o) => o.op === 'Tj' && o.operands[0] === 'Hi');
      expect(fillIndex).toBeGreaterThanOrEqual(0);
      expect(fillIndex).toBeLessThan(tjIndex); // box is drawn before the text on top of it
      expect(colorsBefore(ops, fillIndex + 1).fill).toEqual([1, 0, 0]);
    });

    it('a div with only a border (no background) draws a stroked rect, not a filled one', async () => {
      const { buffer } = await convertPdf('<div style="border:2px solid #00ff00"><p>Hi</p></div>');
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      expect(ops.some((o) => o.op === 'f')).toBe(false);
      const strokeIndex = ops.findIndex((o) => o.op === 'S');
      expect(strokeIndex).toBeGreaterThanOrEqual(0);
      expect(colorsBefore(ops, strokeIndex + 1).stroke).toEqual([0, 1, 0]);
    });

    it('border-radius on a background box draws a curved path (Bézier "c" ops), not a plain rectangle', async () => {
      const { buffer } = await convertPdf('<div style="background:#eeeeee;border-radius:12px"><p>Hi</p></div>');
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      expect(ops.some((o) => o.op === 'c')).toBe(true);
      expect(ops.some((o) => o.op === 're' && o.operands[2] > 400)).toBe(false); // no full-width plain rect
    });

    it('padding insets the text away from the page margin', async () => {
      const withPadding = await convertPdf('<div style="background:#eeeeee;padding:20pt"><p>Hi</p></div>');
      const plain = await convertPdf('<p>Hi</p>');
      const paddedX = getPageContentOps(parsePdf(withPadding.buffer), getPages(parsePdf(withPadding.buffer))[0])
        .find((o) => o.op === 'Tm').operands[4];
      const plainX = getPageContentOps(parsePdf(plain.buffer), getPages(parsePdf(plain.buffer))[0])
        .find((o) => o.op === 'Tm').operands[4];
      expect(paddedX).toBeGreaterThan(plainX);
      expect(paddedX - plainX).toBeCloseTo(20, 0);
    });

    it('box-shadow draws an extra offset rect behind the background/border box', async () => {
      const withShadow = await convertPdf('<div style="border:1px solid #000;box-shadow:0 4px 8px rgba(0,0,0,.2)"><p>Hi</p></div>');
      const withoutShadow = await convertPdf('<div style="border:1px solid #000"><p>Hi</p></div>');
      const shadowRectCount = getPageContentOps(parsePdf(withShadow.buffer), getPages(parsePdf(withShadow.buffer))[0])
        .filter((o) => o.op === 're').length;
      const plainRectCount = getPageContentOps(parsePdf(withoutShadow.buffer), getPages(parsePdf(withoutShadow.buffer))[0])
        .filter((o) => o.op === 're').length;
      expect(shadowRectCount).toBe(plainRectCount + 1);
    });

    it('a multi-paragraph card (cardGroup) shares exactly one drawn box, not one per paragraph', async () => {
      const { buffer } = await convertPdf('<div style="border:1px solid #333"><h2>Title</h2><p>Body one</p><p>Body two</p></div>');
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      // One border box for the whole card, not three (one per paragraph/heading).
      expect(ops.filter((o) => o.op === 'S').length).toBe(1);
      const texts = ops.filter((o) => o.op === 'Tj').map((o) => o.operands[0]);
      expect(texts).toEqual(expect.arrayContaining(['Title', 'Body one', 'Body two']));
    });

    it('shrinkToFit (inline-block) narrows the drawn box to its own content instead of the full page width', async () => {
      const { buffer } = await convertPdf('<p style="display:inline-block;background:#ddd">Short</p>');
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const rect = ops.find((o) => o.op === 're');
      expect(rect.operands[2]).toBeLessThan(468); // less than the full 6in content width at 1in margins
    });

    it('margin-top/margin-bottom widen the gap around a styled box beyond the default block spacing', async () => {
      const { buffer } = await convertPdf('<p>Before</p><div style="background:#eee;margin-top:50pt;margin-bottom:50pt"><p>Boxed</p></div><p>After</p>');
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const beforeY = ops.find((o) => o.op === 'Tj' && o.operands[0] === 'Before');
      const afterY = ops.find((o) => o.op === 'Tj' && o.operands[0] === 'After');
      const beforeTm = [...ops.slice(0, ops.indexOf(beforeY))].reverse().find((o) => o.op === 'Tm');
      const afterTm = [...ops.slice(0, ops.indexOf(afterY))].reverse().find((o) => o.op === 'Tm');
      // "Before"'s own Tm is above "After"'s further down the page by more
      // than the default ~28pt (2 blocks' worth of BLOCK_SPACING_PT + line
      // height) a totally unstyled 3-paragraph run would use.
      expect(beforeTm.operands[5] - afterTm.operands[5]).toBeGreaterThan(28);
    });

    it('adjacent margins collapse to their max, not their sum — a browser never doubles the gap', async () => {
      const tmYFor = (ops, text) => {
        const idx = ops.findIndex((o) => o.op === 'Tj' && o.operands[0] === text);
        return [...ops.slice(0, idx)].reverse().find((o) => o.op === 'Tm').operands[5];
      };
      const gapBetween = async (html) => {
        const { buffer } = await convertPdf(html);
        const ops = getPageContentOps(parsePdf(buffer), getPages(parsePdf(buffer))[0]);
        return tmYFor(ops, 'Heading') - tmYFor(ops, 'Paragraph');
      };

      // margin-bottom:16pt only vs. margin-bottom:16pt + margin-top:16pt —
      // real CSS collapsing makes these identical (max(16,0) === max(16,16)).
      const oneSided = await gapBetween('<h2 style="margin-bottom:16pt">Heading</h2><p style="margin-top:0">Paragraph</p>');
      const bothSided = await gapBetween('<h2 style="margin-bottom:16pt">Heading</h2><p style="margin-top:16pt">Paragraph</p>');
      expect(bothSided).toBeCloseTo(oneSided, 0);

      // The larger of two unequal adjoining margins wins, not their sum.
      const unequal = await gapBetween('<h2 style="margin-bottom:10pt">Heading</h2><p style="margin-top:30pt">Paragraph</p>');
      expect(unequal).toBeCloseTo(oneSided + (30 - 16), 0); // the 30pt side dominates, replacing the 16pt one entirely
    });
  });

  describe('CSS grid/flex layout (layoutGroup)', () => {
    it('a 3-column grid lays its boxes out side by side, not stacked', async () => {
      const html = `
        <div class="grid" style="display:grid;grid-template-columns:repeat(3,1fr)">
          <div style="background:#eef4ff"><h3>Revenue</h3>$2.5M</div>
          <div style="background:#eef4ff"><h3>Customers</h3>12,420</div>
          <div style="background:#eef4ff"><h3>Growth</h3>23%</div>
        </div>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);

      const tmFor = (text) => {
        const idx = ops.findIndex((o) => o.op === 'Tj' && o.operands[0] === text);
        return [...ops.slice(0, idx)].reverse().find((o) => o.op === 'Tm');
      };
      const revenue = tmFor('Revenue');
      const customers = tmFor('Customers');
      const growth = tmFor('Growth');

      // All three headings sit on the same row (same baseline Y)...
      expect(revenue.operands[5]).toBeCloseTo(customers.operands[5], 0);
      expect(customers.operands[5]).toBeCloseTo(growth.operands[5], 0);
      // ...at three strictly increasing X positions (side by side, not stacked).
      expect(customers.operands[4]).toBeGreaterThan(revenue.operands[4]);
      expect(growth.operands[4]).toBeGreaterThan(customers.operands[4]);
    });

    it('a plain display:flex row (no grid-template-columns) also lays cells out side by side', async () => {
      const html = '<div style="display:flex"><p>One</p><p>Two</p><p>Three</p></div>';
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const tmFor = (text) => {
        const idx = ops.findIndex((o) => o.op === 'Tj' && o.operands[0] === text);
        return [...ops.slice(0, idx)].reverse().find((o) => o.op === 'Tm');
      };
      const one = tmFor('One');
      const two = tmFor('Two');
      const three = tmFor('Three');
      expect(one.operands[5]).toBeCloseTo(two.operands[5], 0);
      expect(two.operands[4]).toBeGreaterThan(one.operands[4]);
      expect(three.operands[4]).toBeGreaterThan(two.operands[4]);
    });

    it('a grid cell with more than 3 items wraps into a second row', async () => {
      const html = `
        <div style="display:grid;grid-template-columns:repeat(2,1fr)">
          <p>A</p><p>B</p><p>C</p>
        </div>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);
      const tmFor = (text) => {
        const idx = ops.findIndex((o) => o.op === 'Tj' && o.operands[0] === text);
        return [...ops.slice(0, idx)].reverse().find((o) => o.op === 'Tm');
      };
      const a = tmFor('A');
      const b = tmFor('B');
      const c = tmFor('C');
      // A and B share the first row...
      expect(a.operands[5]).toBeCloseTo(b.operands[5], 0);
      // ...C starts a second row, further down the page.
      expect(c.operands[5]).toBeLessThan(a.operands[5]);
    });

    it('a grid cell whose heading+value both inherit the same box background from one wrapper div draws ONE shared box, not one per node', async () => {
      // `.box{background:#eef4ff}` alone (no border, no explicit color) is
      // exactly the pattern dom-to-idm's own cross-sibling cardGroup
      // heuristic doesn't trust as "one bounded panel" — this cell-level
      // synthesis is what still merges it into one box instead of leaving
      // the heading and value as two separate stacked boxes.
      const html = `
        <div style="display:grid;grid-template-columns:repeat(3,1fr)">
          <div style="background:#eef4ff"><h3>Revenue</h3>$2.5M</div>
          <div style="background:#eef4ff"><h3>Customers</h3>12,420</div>
          <div style="background:#eef4ff"><h3>Growth</h3>23%</div>
        </div>`;
      const { buffer } = await convertPdf(html);
      const pdf = parsePdf(buffer);
      const ops = getPageContentOps(pdf, getPages(pdf)[0]);

      const fillOps = ops.filter((o) => o.op === 'f');
      expect(fillOps).toHaveLength(3); // one filled box per column, not six

      // The single box behind "Revenue" also covers "$2.5M" underneath it —
      // both texts' Y positions fall strictly inside that one rect's span.
      const rect = ops.find((o) => o.op === 're');
      const [rx, ry, , rh] = rect.operands;
      const tmFor = (text) => {
        const idx = ops.findIndex((o) => o.op === 'Tj' && o.operands[0] === text);
        return [...ops.slice(0, idx)].reverse().find((o) => o.op === 'Tm');
      };
      const headingY = tmFor('Revenue').operands[5];
      const valueY = tmFor('$2.5M').operands[5];
      expect(headingY).toBeGreaterThan(ry);
      expect(headingY).toBeLessThan(ry + rh);
      expect(valueY).toBeGreaterThan(ry);
      expect(valueY).toBeLessThan(ry + rh);
    });
  });

  describe('whole-document acceptance', () => {
    it('a document with headings, mixed inline formatting, nested lists, a table, an image, a hyperlink, and enough content to force multiple pages converts consistently end to end', async () => {
      const png = buildTestPng({ width: 4, height: 4, colorType: 0 });
      // 150, not 50: with no invented default gap between blocks (spacing now
      // comes only from each element's own CSS margin, see dom-to-idm.js),
      // plain <p>s with no margin pack far tighter per page than before.
      const filler = Array.from({ length: 150 }, (_, i) => `<p>Filler paragraph ${i}.</p>`).join('');
      const html = `
        <h1>Report Title</h1>
        <p>Some <b>bold</b> and <i>italic</i> text with a <a href="https://example.com">link</a>.</p>
        <ul><li>Item one</li><li>Item two<ul><li>Nested</li></ul></li></ul>
        <table><tr><th>Metric</th><th>Value</th></tr><tr><td>Revenue</td><td>$100</td></tr></table>
        <img src="${dataUri('image/png', png)}" alt="chart">
        ${filler}
      `;
      const { buffer, warnings } = await convertPdf(html, { title: 'Acceptance Test' });
      const pdf = parsePdf(buffer);
      const pages = getPages(pdf);
      expect(pages.length).toBeGreaterThanOrEqual(3);

      // Consecutive same-style words within a line are drawn as one Tj
      // segment (e.g. "Report Title" as a single call, since h1 has no
      // internal style variation) — so this checks substring containment
      // in the joined text rather than exact word-array membership.
      const joinedText = pages
        .flatMap((p) => getPageContentOps(pdf, p).filter((o) => o.op === 'Tj').map((o) => o.operands[0]))
        .join(' ');
      for (const expected of ['Report', 'bold', 'italic', 'Item', 'Revenue', 'Filler']) {
        expect(joinedText).toContain(expected);
      }

      const allAnnots = pages.flatMap((p) => getAnnotations(pdf, p));
      expect(allAnnots.some((a) => a.A?.URI === 'https://example.com')).toBe(true);

      const info = resolveRef(pdf, pdf.trailer.Info);
      expect(info.Title).toBe('Acceptance Test');
      expect(warnings.filter((w) => w.severity === 'warning' && w.code.startsWith('pdf-unsupported'))).toHaveLength(0);
    });
  });
});
