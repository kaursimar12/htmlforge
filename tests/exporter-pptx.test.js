import { describe, it, expect } from 'vitest';
import { deflateSync } from 'node:zlib';
import { convert } from '../src/index.js';
import { readPptxParts } from './helpers/pptx.js';
import { findAll, find, firstChild, attr, textContent } from '../src/utils/xml.js';

// A real (not just claimed) PNG of the given size, for tests that need a
// genuinely oversized intrinsic image rather than a 1x1 fixture — mirrors
// exporter-pdf.test.js's own buildTestPng.
function buildTestPngBytes(width, height) {
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
  'ppt/presentation.xml',
  'ppt/_rels/presentation.xml.rels',
  'ppt/theme/theme1.xml',
  'ppt/slideMasters/slideMaster1.xml',
  'ppt/slideMasters/_rels/slideMaster1.xml.rels',
  'ppt/slideLayouts/slideLayout1.xml',
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
  'ppt/slides/slide1.xml',
  'ppt/slides/_rels/slide1.xml.rels',
];

function slideTexts(slideXml) {
  return findAll(slideXml, 'a:t').map(textContent);
}

// Title/body placeholders have no `a:prstGeom` at all — only a promoted
// badge/box shape does — so this must tolerate that element being absent
// rather than crash calling attr(undefined, ...).
function isRoundRectShape(sp) {
  const geom = find(sp, 'a:prstGeom');
  return !!geom && attr(geom, 'prst') === 'roundRect';
}

function isEllipseShape(sp) {
  const geom = find(sp, 'a:prstGeom');
  return !!geom && attr(geom, 'prst') === 'ellipse';
}

describe('PPTX exporter — skeleton (PP1)', () => {
  it('produces every required OOXML part for a minimal document', async () => {
    const result = await convert('<p>Hello</p>', { format: 'pptx' });
    const { zip } = readPptxParts(result.buffer);
    const entries = new Set(zip.entries());
    for (const part of REQUIRED_PARTS) expect(entries.has(part)).toBe(true);
  });

  it('every XML/rels part is well-formed and parses without throwing', async () => {
    const result = await convert('<h1>T</h1><p>Body</p>', { format: 'pptx' });
    expect(() => readPptxParts(result.buffer)).not.toThrow();
  });

  it('sets the correct top-level mimeType/extension', async () => {
    const result = await convert('<p>x</p>', { format: 'pptx' });
    expect(result.mimeType).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation');
    expect(result.extension).toBe('.pptx');
  });

  it('[Content_Types].xml declares overrides for every OOXML part', async () => {
    const result = await convert('<h1>A</h1><p>x</p><h1>B</h1><p>y</p>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const overrides = findAll(xml['[Content_Types].xml'], 'Override').map((el) => attr(el, 'PartName'));
    expect(overrides).toContain('/ppt/presentation.xml');
    expect(overrides).toContain('/ppt/slideMasters/slideMaster1.xml');
    expect(overrides).toContain('/ppt/slideLayouts/slideLayout1.xml');
    expect(overrides).toContain('/ppt/theme/theme1.xml');
    expect(overrides).toContain('/ppt/slides/slide1.xml');
    expect(overrides).toContain('/ppt/slides/slide2.xml');
  });

  it('every relationship target resolves to a real part in the package', async () => {
    const result = await convert('<h1>A</h1><p>x</p><h1>B</h1><p>y</p>', { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    const entries = new Set(zip.entries());

    function resolve(baseDir, target) {
      const stack = baseDir ? baseDir.split('/') : [];
      for (const part of target.split('/')) {
        if (part === '..') stack.pop();
        else if (part !== '.') stack.push(part);
      }
      return stack.join('/');
    }

    function relsBaseDir(partName) {
      if (partName === '_rels/.rels') return '';
      const idx = partName.lastIndexOf('/_rels/');
      return idx === -1 ? '' : partName.slice(0, idx);
    }

    for (const [partName, doc] of Object.entries(xml)) {
      if (!partName.endsWith('.rels')) continue;
      const basePartDir = relsBaseDir(partName);
      for (const rel of findAll(doc, 'Relationship')) {
        if (attr(rel, 'TargetMode') === 'External') continue;
        const resolved = resolve(basePartDir, attr(rel, 'Target'));
        expect(entries.has(resolved)).toBe(true);
      }
    }
  });

  it('slideMaster references the layout, and presentation references the master + every slide', async () => {
    const result = await convert('<h1>A</h1><p>x</p>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);

    const masterRId = attr(find(xml['ppt/presentation.xml'], 'p:sldMasterId'), 'r:id');
    const presRels = findAll(xml['ppt/_rels/presentation.xml.rels'], 'Relationship');
    const masterRel = presRels.find((r) => attr(r, 'Id') === masterRId);
    expect(attr(masterRel, 'Target')).toBe('slideMasters/slideMaster1.xml');

    const layoutRId = attr(find(xml['ppt/slideMasters/slideMaster1.xml'], 'p:sldLayoutId'), 'r:id');
    const masterRels = findAll(xml['ppt/slideMasters/_rels/slideMaster1.xml.rels'], 'Relationship');
    const layoutRel = masterRels.find((r) => attr(r, 'Id') === layoutRId);
    expect(attr(layoutRel, 'Target')).toBe('../slideLayouts/slideLayout1.xml');
  });
});

describe('PPTX exporter — slide grouping (PP2)', () => {
  it('one <section> per slide when the document uses explicit sections', async () => {
    const html = '<section><h1>One</h1><p>First</p></section><section><h1>Two</h1><p>Second</p></section><section><h1>Three</h1><p>Third</p></section>';
    const result = await convert(html, { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    const slideCount = [...zip.entries()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
    expect(slideCount).toBe(3);
    expect(slideTexts(xml['ppt/slides/slide1.xml'])).toContain('One');
    expect(slideTexts(xml['ppt/slides/slide2.xml'])).toContain('Two');
    expect(slideTexts(xml['ppt/slides/slide3.xml'])).toContain('Three');
  });

  it('falls back to splitting on top-level h1/h2 when there is no <section> at all', async () => {
    const html = '<h1>Slide A</h1><p>a-body</p><h2>Slide B</h2><p>b-body</p>';
    const result = await convert(html, { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    const slideCount = [...zip.entries()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
    expect(slideCount).toBe(2);
    expect(slideTexts(xml['ppt/slides/slide1.xml'])).toEqual(expect.arrayContaining(['Slide A', 'a-body']));
    expect(slideTexts(xml['ppt/slides/slide2.xml'])).toEqual(expect.arrayContaining(['Slide B', 'b-body']));
  });

  it('content before the first heading (no <section>) becomes its own titleless leading slide', async () => {
    const html = '<p>intro</p><h1>Slide A</h1><p>a-body</p>';
    const result = await convert(html, { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    const slideCount = [...zip.entries()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
    expect(slideCount).toBe(2);
    expect(slideTexts(xml['ppt/slides/slide1.xml'])).toEqual(['intro']);
    expect(findAll(xml['ppt/slides/slide1.xml'], 'p:ph').some((ph) => attr(ph, 'type') === 'title')).toBe(false);
  });

  it('a whole document with no headings and no sections becomes one slide', async () => {
    const html = '<p>only paragraph</p>';
    const result = await convert(html, { format: 'pptx' });
    const { zip } = readPptxParts(result.buffer);
    const slideCount = [...zip.entries()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
    expect(slideCount).toBe(1);
  });

  it('a <section> mixed among flat top-level content becomes its own slide, without dropping the rest', async () => {
    const html = '<h1>Intro</h1><p>intro-body</p><section><h1>Sectioned</h1><p>sec-body</p></section><h1>Outro</h1><p>outro-body</p>';
    const result = await convert(html, { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    const slideCount = [...zip.entries()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
    expect(slideCount).toBe(3);
    expect(slideTexts(xml['ppt/slides/slide1.xml'])).toEqual(expect.arrayContaining(['Intro', 'intro-body']));
    expect(slideTexts(xml['ppt/slides/slide2.xml'])).toEqual(expect.arrayContaining(['Sectioned', 'sec-body']));
    expect(slideTexts(xml['ppt/slides/slide3.xml'])).toEqual(expect.arrayContaining(['Outro', 'outro-body']));
  });
});

describe('PPTX exporter — title/body placeholders (PP3)', () => {
  it('first h1/h2 becomes the title placeholder; other text becomes body paragraphs', async () => {
    const result = await convert('<h1>My Title</h1><p>Paragraph one</p><p>Paragraph two</p>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];

    const titleShape = findAll(slide, 'p:sp').find((sp) => attr(find(sp, 'p:ph'), 'type') === 'title');
    expect(textContent(find(titleShape, 'a:t'))).toBe('My Title');

    const bodyShape = findAll(slide, 'p:sp').find((sp) => attr(find(sp, 'p:ph'), 'type') === 'body');
    expect(slideTexts(bodyShape)).toEqual(['Paragraph one', 'Paragraph two']);
  });

  it('a slide with no heading has no title placeholder', async () => {
    const result = await convert('<p>just text</p>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const hasTitlePh = findAll(xml['ppt/slides/slide1.xml'], 'p:ph').some((ph) => attr(ph, 'type') === 'title');
    expect(hasTitlePh).toBe(false);
  });

  it('unstyled h3/h4/h5/h6 in the body get a decreasing default size, not just bold', async () => {
    // Regression: a heading with no CSS font-size previously rendered as
    // merely bold text at the placeholder's own default size — an h3 was
    // visually indistinguishable from an h6, both looking identical to a
    // plain bold paragraph.
    const html = '<h1>T</h1><h3>Three</h3><h4>Four</h4><h5>Five</h5><h6>Six</h6>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const bodyShape = findAll(xml['ppt/slides/slide1.xml'], 'p:sp').find((sp) => attr(find(sp, 'p:ph'), 'type') === 'body');
    const sizeOf = (text) => {
      const run = findAll(bodyShape, 'a:r').find((r) => textContent(find(r, 'a:t')) === text);
      return Number(attr(find(run, 'a:rPr'), 'sz'));
    };
    const sizes = [sizeOf('Three'), sizeOf('Four'), sizeOf('Five'), sizeOf('Six')];
    expect(sizes.every((s) => s > 0)).toBe(true);
    expect(sizes).toEqual([...sizes].sort((a, b) => b - a)); // strictly decreasing h3 > h4 > h5 > h6
  });

  it('an explicit CSS font-size on a heading still overrides the level default', async () => {
    const html = '<h1>T</h1><style>h3{font-size:100px}</style><h3>Big</h3>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const run = findAll(xml['ppt/slides/slide1.xml'], 'a:r').find((r) => textContent(find(r, 'a:t')) === 'Big');
    // 100px -> 75pt (print-accurate) -> x4/3 slide scale -> 100pt -> hundredths
    expect(Number(attr(find(run, 'a:rPr'), 'sz'))).toBe(10000);
  });
});

describe('PPTX exporter — lists (PP4)', () => {
  it('unordered list items get a:buChar bullets', async () => {
    const result = await convert('<h1>T</h1><ul><li>First</li><li>Second</li></ul>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const bullets = findAll(slide, 'a:buChar');
    expect(bullets.length).toBe(2);
    expect(slideTexts(slide)).toEqual(expect.arrayContaining(['First', 'Second']));
  });

  it('ordered list items get a:buAutoNum, honoring a non-default start', async () => {
    const result = await convert('<h1>T</h1><ol start="3"><li>Third</li><li>Fourth</li></ol>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const autoNums = findAll(slide, 'a:buAutoNum');
    expect(autoNums.length).toBe(2);
    expect(attr(autoNums[0], 'startAt')).toBe('3');
  });

  it('nested list items increase a:pPr lvl', async () => {
    const html = '<h1>T</h1><ul><li>Top<ul><li>Nested</li></ul></li></ul>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const paragraphs = findAll(slide, 'a:p').filter((p) => findAll(p, 'a:buChar').length > 0);
    const levels = paragraphs.map((p) => Number(attr(find(p, 'a:pPr'), 'lvl') ?? 0));
    expect(levels).toEqual([0, 1]);
  });

  it('plain (non-list) paragraphs explicitly suppress bullets with a:buNone', async () => {
    const result = await convert('<h1>T</h1><p>Not a bullet</p>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const bodyShape = findAll(slide, 'p:sp').find((sp) => attr(find(sp, 'p:ph'), 'type') === 'body');
    expect(findAll(bodyShape, 'a:buNone').length).toBeGreaterThan(0);
  });
});

describe('PPTX exporter — inline run formatting', () => {
  it('<mark> renders as a real text highlight, not just plain text', async () => {
    const result = await convert('<h1>T</h1><p><mark>Highlighted</mark></p>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const run = findAll(xml['ppt/slides/slide1.xml'], 'a:r').find((r) => textContent(find(r, 'a:t')) === 'Highlighted');
    expect(find(run, 'a:highlight')).toBeTruthy();
  });

  it('<sub>/<sup> set a positive/negative a:rPr baseline', async () => {
    const result = await convert('<h1>T</h1><p>H<sub>2</sub>O, X<sup>2</sup></p>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const runs = findAll(xml['ppt/slides/slide1.xml'], 'a:r');
    const sub = runs.find((r) => textContent(find(r, 'a:t')) === '2' && Number(attr(find(r, 'a:rPr'), 'baseline')) < 0);
    const sup = runs.find((r) => textContent(find(r, 'a:t')) === '2' && Number(attr(find(r, 'a:rPr'), 'baseline')) > 0);
    expect(sub).toBeTruthy();
    expect(sup).toBeTruthy();
  });
});

describe('PPTX exporter — images (PP5)', () => {
  const PNG_1X1 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

  it('embeds the image under ppt/media and links it via a per-slide relationship', async () => {
    const html = `<h1>T</h1><img src="${PNG_1X1}" width="200" height="100" alt="tiny"/>`;
    const result = await convert(html, { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    expect(zip.entries()).toContain('ppt/media/image1.png');

    const slide = xml['ppt/slides/slide1.xml'];
    const blip = find(slide, 'a:blip');
    const relId = attr(blip, 'r:embed');
    const rel = findAll(xml['ppt/slides/_rels/slide1.xml.rels'], 'Relationship').find((r) => attr(r, 'Id') === relId);
    expect(attr(rel, 'Target')).toBe('../media/image1.png');
  });

  it('an unresolvable image renders no <p:pic> and does not crash', async () => {
    const html = '<h1>T</h1><img src="not-a-real-path-xyz.png" alt="broken"/>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    expect(findAll(xml['ppt/slides/slide1.xml'], 'p:pic').length).toBe(0);
  });

  it('a declared width/height attribute is ignored in favor of the image\'s own real intrinsic size (1x1)', async () => {
    // PNG_1X1 is a real, decodable 1x1 image — a stale width="200"
    // height="100" attribute must not override what the file actually is.
    const html = `<h1>T</h1><img src="${PNG_1X1}" width="200" height="100" alt="tiny"/>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const pic = find(xml['ppt/slides/slide1.xml'], 'p:pic');
    const ext = find(pic, 'a:ext');
    expect(Number(attr(ext, 'cx'))).toBe(9525); // 1px * 9525 EMU/px, not 200px's 1905000
    expect(Number(attr(ext, 'cy'))).toBe(9525);
  });
});

describe('PPTX exporter — SVG vector shapes', () => {
  it('a rect+circle logo draws two separately-positioned shapes, the circle centered within the rect', async () => {
    const html = '<svg width="200" height="120"><rect width="200" height="120" fill="#2563eb"/><circle cx="100" cy="60" r="35" fill="orange"/></svg>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const shapes = findAll(xml['ppt/slides/slide1.xml'], 'p:sp');
    expect(shapes).toHaveLength(2);

    const [rectShape, circleShape] = shapes;
    expect(attr(find(rectShape, 'a:prstGeom'), 'prst')).toBe('rect');
    expect(attr(find(rectShape, 'a:srgbClr'), 'val')).toBe('2563EB');
    expect(attr(find(circleShape, 'a:prstGeom'), 'prst')).toBe('ellipse');
    expect(attr(find(circleShape, 'a:srgbClr'), 'val')).toBe('FFA500');

    const rectOff = find(rectShape, 'a:off');
    const rectExt = find(rectShape, 'a:ext');
    const circleOff = find(circleShape, 'a:off');
    const circleExt = find(circleShape, 'a:ext');
    const rectCenterX = Number(attr(rectOff, 'x')) + Number(attr(rectExt, 'cx')) / 2;
    const circleCenterX = Number(attr(circleOff, 'x')) + Number(attr(circleExt, 'cx')) / 2;
    expect(circleCenterX).toBeCloseTo(rectCenterX, 0);
  });

  it('a viewBox rescales shape geometry into the declared display size', async () => {
    // viewBox is a 24x24 icon space shown at 48x48 (2x) — a 20x20 rect
    // becomes 40x40 raw units, then 40 * 9525 EMU/px = 381000 EMU.
    const html = '<svg width="48" height="48" viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="20" fill="red"/></svg>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const shape = find(xml['ppt/slides/slide1.xml'], 'p:sp');
    const ext = find(shape, 'a:ext');
    expect(Number(attr(ext, 'cx'))).toBe(381000);
    expect(Number(attr(ext, 'cy'))).toBe(381000);
  });

  it('a polygon renders as a real a:custGeom closed path with the right fill/stroke', async () => {
    const html = '<svg width="100" height="100"><polygon points="0,0 10,0 5,10" fill="red" stroke="blue"/></svg>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const shape = find(xml['ppt/slides/slide1.xml'], 'p:sp');
    const custGeom = find(shape, 'a:custGeom');
    expect(custGeom).toBeTruthy();
    expect(findAll(custGeom, 'a:lnTo')).toHaveLength(2);
    expect(find(custGeom, 'a:close')).toBeTruthy();
    const fills = findAll(shape, 'a:srgbClr').map((el) => attr(el, 'val'));
    expect(fills).toEqual(['FF0000', '0000FF']);
  });

  it('a line renders as a stroked, unfilled line-preset shape', async () => {
    const html = '<svg width="100" height="100"><line x1="0" y1="0" x2="100" y2="100" stroke="black"/></svg>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const shape = find(xml['ppt/slides/slide1.xml'], 'p:sp');
    expect(attr(find(shape, 'a:prstGeom'), 'prst')).toBe('line');
    expect(find(shape, 'a:noFill')).toBeTruthy();
    expect(find(shape, 'a:ln')).toBeTruthy();
  });

  it('an SVG with only unsupported content (e.g. a <path>) still drops silently', async () => {
    const html = '<svg><path d="M10 10 L90 90"/></svg><p>after</p>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    expect(findAll(xml['ppt/slides/slide1.xml'], 'p:sp').length).toBeLessThanOrEqual(1); // only the text placeholder, no shape
    expect(textContent(xml['ppt/slides/slide1.xml'])).toContain('after');
  });
});

describe('PPTX exporter — tables (PP6)', () => {
  it('renders a table as a p:graphicFrame/a:tbl with the right grid size', async () => {
    const html = '<h1>T</h1><table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const frame = find(slide, 'p:graphicFrame');
    expect(frame).toBeTruthy();
    expect(findAll(frame, 'a:gridCol').length).toBe(2);
    expect(findAll(frame, 'a:tr').length).toBe(2);
    expect(slideTexts(frame)).toEqual(['a', 'b', 'c', 'd']);
  });

  it('bolds header row cells from <thead>/<th>', async () => {
    const html = '<h1>T</h1><table><thead><tr><th>H1</th></tr></thead><tbody><tr><td>v1</td></tr></tbody></table>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    const rows = findAll(frame, 'a:tr');
    expect(attr(find(rows[0], 'a:rPr'), 'b')).toBe('1');
    expect(attr(find(rows[1], 'a:rPr'), 'b')).toBeUndefined();
  });

  it('rowspan produces rowSpan on the origin cell and vMerge on the continuation', async () => {
    const html = '<h1>T</h1><table><tr><td rowspan="2">R</td><td>b1</td></tr><tr><td>b2</td></tr></table>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    const rows = findAll(frame, 'a:tr');
    expect(attr(findAll(rows[0], 'a:tc')[0], 'rowSpan')).toBe('2');
    expect(attr(findAll(rows[1], 'a:tc')[0], 'vMerge')).toBe('1');
  });

  it('colspan produces gridSpan on the origin cell and an hMerge placeholder', async () => {
    const html = '<h1>T</h1><table><tr><td colspan="2">wide</td></tr><tr><td>a</td><td>b</td></tr></table>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    const rows = findAll(frame, 'a:tr');
    const firstRowCells = findAll(rows[0], 'a:tc');
    expect(attr(firstRowCells[0], 'gridSpan')).toBe('2');
    expect(attr(firstRowCells[1], 'hMerge')).toBe('1');
    // Every row must still carry as many <a:tc> as there are grid columns.
    expect(firstRowCells.length).toBe(findAll(frame, 'a:gridCol').length);
  });

  it('a <th> background/color cascades onto the cell fill and run color', async () => {
    const html = '<h1>T</h1><style>th{background:#2563EB;color:white;}</style><table><tr><th>Name</th></tr><tr><td>x</td></tr></table>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    const headerCell = findAll(frame, 'a:tc')[0];
    expect(attr(find(firstChild(find(headerCell, 'a:tcPr'), 'a:solidFill'), 'a:srgbClr'), 'val')).toBe('2563EB');
    expect(attr(find(find(headerCell, 'a:rPr'), 'a:srgbClr'), 'val')).toBe('FFFFFF');
  });

  it('a cell with text-align:center gets algn="ctr" on its a:pPr', async () => {
    const html = '<h1>T</h1><table><tr><td style="text-align:center">mid</td></tr></table>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    expect(attr(find(frame, 'a:pPr'), 'algn')).toBe('ctr');
  });

  it('a colored colspan cell fills its hMerge placeholder to match', async () => {
    const html = '<h1>T</h1><table><tr><td colspan="2" style="background:#dddddd">wide</td></tr><tr><td>a</td><td>b</td></tr></table>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    const firstRowCells = findAll(findAll(frame, 'a:tr')[0], 'a:tc');
    const fillOf = (tc) => attr(find(firstChild(find(tc, 'a:tcPr'), 'a:solidFill'), 'a:srgbClr'), 'val');
    expect(fillOf(firstRowCells[0])).toBe('DDDDDD');
    expect(fillOf(firstRowCells[1])).toBe('DDDDDD');
  });

  // A native DrawingML table cell can't host its own separate <p:pic> shape
  // (see tcPrXml's own doc comment in shapes.js) — an image inside a <td>
  // used to silently vanish (fell into realTableCellXml's plain-text branch,
  // where renderRun's switch has no 'image' case) — the same "Preview"
  // column an image-comparison table's own photo cell needs.
  const PNG_1X1 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

  it('an <img> alone in a table cell becomes a real floating picture positioned over that cell, not an empty paragraph', async () => {
    const html = `<h1>T</h1><table><tr><td>Label</td><td><img src="${PNG_1X1}" alt="x"></td></tr></table>`;
    const result = await convert(html, { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    expect(zip.entries()).toContain('ppt/media/image1.png');
    const slide = xml['ppt/slides/slide1.xml'];

    // A native a:tc can't nest a p:pic — the picture is its own floating
    // shape, a sibling of the table's own p:graphicFrame in the slide's
    // shape tree, positioned exactly over the cell it belongs to.
    const pic = find(slide, 'p:pic');
    expect(pic).toBeTruthy();
    const blip = find(pic, 'a:blip');
    const relId = attr(blip, 'r:embed');
    const rel = findAll(xml['ppt/slides/_rels/slide1.xml.rels'], 'Relationship').find((r) => attr(r, 'Id') === relId);
    expect(attr(rel, 'Target')).toBe('../media/image1.png');

    // The cell itself still renders as an ordinary (empty) text cell — no
    // blipFill/image reference left on its own a:tcPr.
    const frame = find(slide, 'p:graphicFrame');
    const cells = findAll(frame, 'a:tc');
    expect(find(find(cells[1], 'a:tcPr'), 'a:blip')).toBeFalsy();

    // Positioned to the right of the first column (over the second,
    // image-holding cell), and within the table's own vertical extent.
    const frameOff = find(find(frame, 'p:xfrm'), 'a:off');
    const picOff = find(find(pic, 'p:spPr'), 'a:off');
    expect(Number(attr(picOff, 'x'))).toBeGreaterThan(Number(attr(frameOff, 'x')));
    expect(Number(attr(picOff, 'y'))).toBeGreaterThanOrEqual(Number(attr(frameOff, 'y')));
  });

  it('a table cell mixing text and an <img> keeps the text AND draws the image as its own floating picture', async () => {
    const html = `<h1>T</h1><table><tr><td><b>Caption</b><img src="${PNG_1X1}" alt="x"></td></tr></table>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const frame = find(slide, 'p:graphicFrame');
    const cell = find(frame, 'a:tc');
    expect(slideTexts(cell)).toContain('Caption');
    expect(find(slide, 'p:pic')).toBeTruthy();
  });

  it('a cell holding only an image gets a row tall enough for the image\'s own real size, not squeezed to a single text line', async () => {
    // PNG_1X1 has no declared width/height, so resolveImagePxSize falls
    // back to its real intrinsic size — 1x1 px, trivially tiny — so this
    // uses a fixture with a genuinely larger real size instead, matching a
    // realistic "photo preview" cell.
    const png300x150 = `data:image/png;base64,${buildTestPngBytes(300, 150).toString('base64')}`;
    const html = `<h1>T</h1><table><tr><td style="width:20%">Label</td><td style="width:20%"><img src="${png300x150}"></td></tr></table>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    const rowHeight = Number(attr(find(frame, 'a:tr'), 'h'));
    // A single text line at the default 18pt cell font is roughly
    // 18*1.2*12700 ≈ 274320 EMU plus padding — a 300x150 photo scaled to
    // fit a narrow 20%-of-slide column needs meaningfully more than that.
    expect(rowHeight).toBeGreaterThan(400000);
  });

  // rowNaturalHeightEmu's own row-height estimate and buildTableGraphicFrame's
  // own floating-picture box must agree on the exact same available width —
  // using the row's own declared height for the row but the FULL (unpadded)
  // column width for the picture's box let the picture's own aspect-fit
  // scale up past what the row actually budgeted for, so the picture's real
  // drawn bottom edge silently exceeded the row's own bottom edge in real
  // PowerPoint (an image "spilling out of the table").
  it('a cell-image\'s own floating picture never extends past its row\'s own top/bottom edge', async () => {
    const png300x150 = `data:image/png;base64,${buildTestPngBytes(300, 150).toString('base64')}`;
    const html = `<h1>T</h1><table><tr><td style="width:20%">Label</td><td style="width:20%"><img src="${png300x150}"></td></tr></table>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    const frameOff = find(find(frame, 'p:xfrm'), 'a:off');
    const rowHeight = Number(attr(find(frame, 'a:tr'), 'h'));
    const rowTop = Number(attr(frameOff, 'y'));
    const rowBottom = rowTop + rowHeight;

    const pic = find(xml['ppt/slides/slide1.xml'], 'p:pic');
    const picOff = find(find(pic, 'p:spPr'), 'a:off');
    const picExt = find(find(pic, 'p:spPr'), 'a:ext');
    const picTop = Number(attr(picOff, 'y'));
    const picBottom = picTop + Number(attr(picExt, 'cy'));

    expect(picTop).toBeGreaterThanOrEqual(rowTop);
    expect(picBottom).toBeLessThanOrEqual(rowBottom);
  });

  // A row's declared height is only an ESTIMATE (this file has no real text
  // layout engine) — if a cell holds a long, unbreakable single token (a
  // bare URL/file path with no <p> wrapper and no spaces to wrap at, e.g.
  // this project's own "Image Source" column), real PowerPoint still wraps
  // it across several rows once it doesn't fit the column, even though
  // estimateWrappedLines' own word-only contract can't split it. Under-
  // estimating that row silently pushed every shape positioned after this
  // table (this project's own header logo cell, its blipFill included)
  // down into a table whose REAL rendered height was taller than what this
  // exporter reserved, producing a visible overlap in real PowerPoint.
  it('a cell holding one long unbreakable token (no spaces) gets a taller row than a short cell, not the same flat height', async () => {
    const longUrl = 'https://example.com/a/very/long/path/that/keeps/going/and/going/image.png';
    const html = `<h1>T</h1><table><tr><td style="width:20%">short</td><td style="width:20%">${longUrl}</td></tr></table>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const frame = find(xml['ppt/slides/slide1.xml'], 'p:graphicFrame');
    const rowHeight = Number(attr(find(frame, 'a:tr'), 'h'));
    const shortOnlyHtml = '<h1>T</h1><table><tr><td style="width:20%">short</td><td style="width:20%">also short</td></tr></table>';
    const shortResult = await convert(shortOnlyHtml, { format: 'pptx' });
    const shortRowHeight = Number(attr(find(readPptxParts(shortResult.buffer).xml['ppt/slides/slide1.xml'], 'a:tr'), 'h'));
    expect(rowHeight).toBeGreaterThan(shortRowHeight);
  });
});

describe('PPTX exporter — paragraph/heading alignment and code block colors', () => {
  it('text-align:center/right/justify map to a:pPr algn', async () => {
    const html = '<h1>T</h1><p style="text-align:center">c</p><p style="text-align:right">r</p><p style="text-align:justify">j</p>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const bodyShape = findAll(xml['ppt/slides/slide1.xml'], 'p:sp').find((sp) => attr(find(sp, 'p:ph'), 'type') === 'body');
    const aligns = findAll(bodyShape, 'a:pPr').map((p) => attr(p, 'algn'));
    expect(aligns).toEqual(['ctr', 'r', 'just']);
  });

  it('a centered h1 title placeholder also gets algn="ctr"', async () => {
    const result = await convert('<h1 style="text-align:center">Centered</h1>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const titleShape = findAll(xml['ppt/slides/slide1.xml'], 'p:sp').find((sp) => attr(find(sp, 'p:ph'), 'type') === 'title');
    expect(attr(find(titleShape, 'a:pPr'), 'algn')).toBe('ctr');
  });

  it('a <pre> with its own background/color applies a run highlight + fill color', async () => {
    const html = '<h1>T</h1><pre style="background:#272822;color:#eee">line one</pre>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const bodyShape = findAll(xml['ppt/slides/slide1.xml'], 'p:sp').find((sp) => attr(find(sp, 'p:ph'), 'type') === 'body');
    const run = find(bodyShape, 'a:r');
    expect(attr(find(find(run, 'a:highlight'), 'a:srgbClr'), 'val')).toBe('272822');
    expect(attr(find(find(run, 'a:solidFill'), 'a:srgbClr'), 'val')).toBe('EEEEEE');
  });

  it('a long <div> with its own background does not leave same-colored text invisible', async () => {
    // Regression: a div's own `background` stamps `backgroundColor` on the
    // paragraph it wraps (block-level), entirely separate from the run-level
    // styledSpan its own `color` produces — previously only the white text
    // color was applied, with nothing behind it, making it disappear against
    // the slide's own white background. Long enough that it stays plain
    // body-placeholder text (see the badge-promotion test below for the
    // short-text case, which gets its own shape instead).
    const html = '<h1>T</h1><div style="background:#2563EB;color:white">Gradient Background Example Text Here, definitely long enough</div>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const bodyShape = findAll(xml['ppt/slides/slide1.xml'], 'p:sp').find((sp) => attr(find(sp, 'p:ph'), 'type') === 'body');
    const run = findAll(bodyShape, 'a:r').find((r) => textContent(find(r, 'a:t')).includes('Gradient'));
    expect(attr(find(find(run, 'a:highlight'), 'a:srgbClr'), 'val')).toBe('2563EB');
    expect(attr(find(find(run, 'a:solidFill'), 'a:srgbClr'), 'val')).toBe('FFFFFF');
  });
});

describe('PPTX exporter — badge/pill shape promotion', () => {
  it('a short background block (e.g. a badge/circle) becomes its own rounded, filled shape, not a highlighted run', async () => {
    // `display:inline-block` (or an explicit width — see the width-based
    // test below) is what actually signals "shrink to content" in real CSS;
    // a plain `<div>` with no such signal is block-level and defaults to
    // filling its container's width instead (see the full-width-box test
    // in the card shape promotion suite below).
    const html = '<h1>T</h1><div class="circle" style="display:inline-block;background:#2563EB;color:white">72%</div>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const boxShape = findAll(slide, 'p:sp').find((sp) => isRoundRectShape(sp));
    expect(boxShape).toBeTruthy();
    expect(attr(find(find(boxShape, 'p:spPr'), 'a:srgbClr'), 'val')).toBe('2563EB');
    const run = find(boxShape, 'a:r');
    expect(textContent(find(run, 'a:t'))).toBe('72%');
    expect(find(run, 'a:highlight')).toBeUndefined(); // fill is on the shape now, not the run
    expect(attr(find(find(run, 'a:solidFill'), 'a:srgbClr'), 'val')).toBe('FFFFFF');
  });

  it('the badge shape is sized to its own text, not stretched to the content width', async () => {
    const html = '<h1>T</h1><p class="badge" style="display:inline-block;background:#2563EB;color:white">VERSION 1.0</p>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const boxShape = findAll(slide, 'p:sp').find((sp) => isRoundRectShape(sp));
    const ext = find(boxShape, 'a:ext');
    // Content width is ~11.28M EMU at widescreen — a two-word badge must be
    // dramatically narrower than that, not stretched full-width.
    expect(Number(attr(ext, 'cx'))).toBeLessThan(3000000);
  });

  it('an explicit fixed width is honored directly instead of the char-count size guess', async () => {
    // border-radius:50% on an equal width/height box is a circle — see the
    // dedicated ellipse tests below — so this is now found via isEllipseShape,
    // not isRoundRectShape; the point of this test is the size, not the shape.
    const html = '<h1>T</h1><div style="width:120px;height:120px;border-radius:50%;background:#2563EB;color:white">72%</div>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const boxShape = findAll(slide, 'p:sp').find((sp) => isEllipseShape(sp));
    const ext = find(boxShape, 'a:ext');
    expect(Number(attr(ext, 'cx'))).toBe(120 * 9525);
    expect(Number(attr(ext, 'cy'))).toBe(120 * 9525);
  });

  it('border-radius:50% on a square badge renders as a real ellipse, not a rounded rectangle', async () => {
    const html = '<h1>T</h1><div style="width:120px;height:120px;border-radius:50%;background:#2563EB;color:white">72%</div>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    expect(findAll(slide, 'p:sp').some((sp) => isEllipseShape(sp))).toBe(true);
    expect(findAll(slide, 'p:sp').some((sp) => isRoundRectShape(sp))).toBe(false);
  });

  it('border-radius:50% on a non-square box renders as an oval (still ellipse), sized to its own box', async () => {
    const html = '<h1>T</h1><div style="width:200px;height:100px;border-radius:50%;background:#2563EB;color:white">Oval</div>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const ovalShape = findAll(slide, 'p:sp').find((sp) => isEllipseShape(sp));
    expect(ovalShape).toBeTruthy();
    const ext = find(ovalShape, 'a:ext');
    expect(Number(attr(ext, 'cx'))).toBe(200 * 9525);
    expect(Number(attr(ext, 'cy'))).toBe(100 * 9525);
  });

  it('a plain rounded corner (border-radius under 50%) still renders as roundRect, not ellipse', async () => {
    const html = '<h1>T</h1><div style="border:1px solid #ddd;padding:20px;border-radius:10px">Card</div>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    expect(findAll(slide, 'p:sp').some((sp) => isRoundRectShape(sp))).toBe(true);
    expect(findAll(slide, 'p:sp').some((sp) => isEllipseShape(sp))).toBe(false);
  });

  it('a long paragraph with the same background is not promoted to a badge shape', async () => {
    const html = '<h1>T</h1><p style="background:#2563EB;color:white">This is a much longer paragraph that should not become a tiny shrink-wrapped badge shape.</p>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const hasBoxShape = findAll(xml['ppt/slides/slide1.xml'], 'p:sp').some((sp) => isRoundRectShape(sp));
    expect(hasBoxShape).toBe(false);
  });

  it('a linear-gradient background renders as a real DrawingML gradient fill, not a flat color', async () => {
    const html = '<h1>T</h1><p class="badge" style="display:inline-block;padding:10px;background:linear-gradient(90deg,#3B82F6,#9333EA)">Gradient Button</p>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const boxShape = findAll(xml['ppt/slides/slide1.xml'], 'p:sp').find((sp) => find(find(sp, 'p:spPr'), 'a:gradFill'));
    expect(boxShape).toBeTruthy();
    const gradFill = find(find(boxShape, 'p:spPr'), 'a:gradFill');
    const stops = findAll(gradFill, 'a:gs');
    expect(stops.map((gs) => attr(gs, 'pos'))).toEqual(['0', '100000']);
    expect(stops.map((gs) => attr(find(gs, 'a:srgbClr'), 'val'))).toEqual(['3B82F6', '9333EA']);
  });
});

describe('PPTX exporter — full-width box promotion', () => {
  it('a short background/border block with no shrink-to-fit signal becomes a full-width box, not a shrink-wrapped badge', async () => {
    // Regression: a plain block-level <div> (no display:inline-block, no
    // explicit width) with a short line of text and its own background —
    // e.g. a full-width gradient CTA button/banner sitting under a card of
    // the same width — previously got wrongly shrink-wrapped into a tiny
    // pill purely because its text happened to be short.
    const html = '<h1>T</h1><div style="padding:10px;background:linear-gradient(90deg,#3B82F6,#9333EA);color:white">Gradient Background Example</div>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const boxShape = findAll(slide, 'p:sp').find((sp) => find(find(sp, 'p:spPr'), 'a:gradFill'));
    expect(boxShape).toBeTruthy();
    const ext = find(boxShape, 'a:ext');
    // Full content width (~11.28M EMU at widescreen), not shrink-wrapped.
    expect(Number(attr(ext, 'cx'))).toBeGreaterThan(10000000);
  });

  it('a box-shadow renders as a real DrawingML outer shadow effect', async () => {
    const html = '<h1>T</h1><div style="border:1px solid #ddd;box-shadow:0 3px 10px rgba(0,0,0,.1)"><p>Card line</p></div>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const cardShape = findAll(slide, 'p:sp').find((sp) => isRoundRectShape(sp));
    expect(cardShape).toBeTruthy();
    expect(find(find(cardShape, 'p:spPr'), 'a:effectLst')).toBeTruthy();
  });

  it('a plain background/shadow on a big page-level wrapper does not turn every short heading inside it into its own shadowed box', async () => {
    // Regression: `.slide{background:white;box-shadow:...}` wrapping a
    // whole section's worth of content (a heading, a badge, a plain
    // paragraph) bled its background onto every plain, otherwise-unstyled
    // heading inside it — and since a section heading is usually short
    // text, it got wrongly promoted into its own full-width "card" shape
    // carrying the wrapper's shadow, purely for sitting inside something
    // with a background.
    const html = `<div style="background:white;box-shadow:0 5px 15px rgba(0,0,0,.15);padding:20px">
      <h2>Section</h2>
      <h3>Short Heading</h3>
      <p>Some body text.</p>
    </div>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    expect(slideTexts(slide)).toContain('Short Heading');
    expect(findAll(slide, 'p:sp').some((sp) => find(find(sp, 'p:spPr'), 'a:effectLst'))).toBe(false);
    expect(findAll(slide, 'p:sp').some((sp) => isRoundRectShape(sp))).toBe(false);
  });

  it('a badge that already declares its own background/inline-block does not also pick up an unrelated ancestor wrapper\'s shadow', async () => {
    const html = `<div style="background:white;box-shadow:0 5px 15px rgba(0,0,0,.15);padding:20px">
      <h1>T</h1>
      <p style="display:inline-block;background:#2563EB;color:white">VERSION 1.0</p>
    </div>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const badgeShape = findAll(slide, 'p:sp').find((sp) => isRoundRectShape(sp));
    expect(badgeShape).toBeTruthy();
    expect(find(find(badgeShape, 'p:spPr'), 'a:effectLst')).toBeUndefined();
  });
});

describe('PPTX exporter — card shape promotion', () => {
  it('a bordered <div> wrapping several blocks becomes one shape behind all of them, not per-paragraph highlights', async () => {
    const html = `<h1>T</h1><div style="border:1px solid #ddd;padding:20px;border-radius:10px">
      <h3>User Card</h3>
      <p>Name: John Doe</p>
      <p>Role: Software Engineer</p>
    </div>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const cardShape = findAll(slide, 'p:sp').find((sp) => isRoundRectShape(sp));
    expect(cardShape).toBeTruthy();
    expect(attr(find(find(cardShape, 'p:spPr'), 'a:ln'), 'w')).toBeTruthy();
    const texts = slideTexts(cardShape);
    expect(texts).toEqual(['User Card', 'Name: John Doe', 'Role: Software Engineer']);
    // Exactly one shape holds all three blocks — not three separately
    // highlighted/bordered runs scattered across the body placeholder.
    expect(findAll(slide, 'p:sp').filter((sp) => isRoundRectShape(sp))).toHaveLength(1);
  });

  it('an already-styled standalone box inside an outer backgrounded container stays its own separate shape, not merged into the outer card', async () => {
    // Regression: an outer wrapper's own background (e.g. a "slide" div with
    // background:white) must not sweep an already self-styled child (a
    // gradient button that already fully styled itself) into the *same*
    // group as an unrelated sibling box — each should keep rendering as its
    // own shape, not merge into one multi-node card spanning both.
    const html = `<h1>T</h1><div style="background:#ffffff;padding:20px">
      <div style="border:1px solid #ddd;padding:10px"><p>Card line</p></div>
      <div style="padding:10px;background:linear-gradient(90deg,#3B82F6,#9333EA)">Go</div>
    </div>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const gradientShape = findAll(slide, 'p:sp').find((sp) => find(find(sp, 'p:spPr'), 'a:gradFill'));
    expect(gradientShape).toBeTruthy();
    // The gradient shape's own text is just "Go" — if it had been wrongly
    // merged with the bordered sibling, "Card line" would appear in the same
    // shape's text runs too.
    expect(slideTexts(gradientShape)).toEqual(['Go']);
    const borderedShape = findAll(slide, 'p:sp').find((sp) => find(find(sp, 'p:spPr'), 'a:ln'));
    expect(borderedShape).toBeTruthy();
    expect(slideTexts(borderedShape)).toEqual(['Card line']);
  });
});

describe('PPTX exporter — flex/grid layout', () => {
  it('a display:flex container lays its children out side by side in one row, not stacked', async () => {
    const html = `<h1>T</h1><div style="display:flex;gap:20px">
      <div style="background:#ef4444;color:white">Box 1</div>
      <div style="background:#10b981;color:white">Box 2</div>
      <div style="background:#f59e0b;color:white">Box 3</div>
    </div>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const shapes = findAll(slide, 'p:sp').filter((sp) => slideTexts(sp).some((t) => /^Box \d$/.test(t)));
    expect(shapes).toHaveLength(3);
    const positions = shapes.map((sp) => {
      const off = find(sp, 'a:off');
      return { x: Number(attr(off, 'x')), y: Number(attr(off, 'y')) };
    });
    // Same row (equal y), strictly increasing x — side by side, not stacked.
    expect(new Set(positions.map((p) => p.y)).size).toBe(1);
    expect(positions[1].x).toBeGreaterThan(positions[0].x);
    expect(positions[2].x).toBeGreaterThan(positions[1].x);
  });

  it('a display:grid container with grid-template-columns wraps cells into that many columns per row', async () => {
    const html = `<h1>T</h1><div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
      <div style="border:1px solid #ddd">Item 1</div>
      <div style="border:1px solid #ddd">Item 2</div>
      <div style="border:1px solid #ddd">Item 3</div>
      <div style="border:1px solid #ddd">Item 4</div>
    </div>`;
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const slide = xml['ppt/slides/slide1.xml'];
    const shapes = findAll(slide, 'p:sp').filter((sp) => slideTexts(sp).some((t) => /^Item \d$/.test(t)));
    expect(shapes).toHaveLength(4);
    const ys = shapes.map((sp) => Number(attr(find(sp, 'a:off'), 'y')));
    // Two distinct rows (2 columns × 2 rows for 4 items), each row shared by 2 cells.
    const rowYs = [...new Set(ys)];
    expect(rowYs).toHaveLength(2);
    expect(ys.filter((y) => y === rowYs[0])).toHaveLength(2);
  });
});

describe('PPTX exporter — slide size option', () => {
  it('defaults to 16:9 widescreen', async () => {
    const result = await convert('<p>x</p>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const sldSz = find(xml['ppt/presentation.xml'], 'p:sldSz');
    expect(attr(sldSz, 'cx')).toBe('12192000');
    expect(attr(sldSz, 'cy')).toBe('6858000');
  });

  it('honors slideSize: "standard" for a 4:3 canvas', async () => {
    const result = await convert('<p>x</p>', { format: 'pptx', slideSize: 'standard' });
    const { xml } = readPptxParts(result.buffer);
    const sldSz = find(xml['ppt/presentation.xml'], 'p:sldSz');
    expect(attr(sldSz, 'cx')).toBe('9144000');
    expect(attr(sldSz, 'cy')).toBe('6858000');
  });
});

describe('PPTX exporter — deck-wide default font', () => {
  it('a <body> font-family becomes the theme\'s default typeface', async () => {
    const html = '<html><head><style>body{font-family:Arial}</style></head><body><h1>T</h1></body></html>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const majorLatin = find(find(xml['ppt/theme/theme1.xml'], 'a:majorFont'), 'a:latin');
    const minorLatin = find(find(xml['ppt/theme/theme1.xml'], 'a:minorFont'), 'a:latin');
    expect(attr(majorLatin, 'typeface')).toBe('Arial');
    expect(attr(minorLatin, 'typeface')).toBe('Arial');
  });

  it('defaults to Calibri when the document declares no body font-family', async () => {
    const result = await convert('<h1>T</h1>', { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const majorLatin = find(find(xml['ppt/theme/theme1.xml'], 'a:majorFont'), 'a:latin');
    expect(attr(majorLatin, 'typeface')).toBe('Calibri');
  });

  it('an element\'s own explicit font-family still overrides the deck default', async () => {
    const html = '<html><head><style>body{font-family:Arial}</style></head><body><h1>T</h1><p style="font-family:Georgia">g</p></body></html>';
    const result = await convert(html, { format: 'pptx' });
    const { xml } = readPptxParts(result.buffer);
    const run = findAll(xml['ppt/slides/slide1.xml'], 'a:r').find((r) => textContent(find(r, 'a:t')) === 'g');
    expect(attr(find(run, 'a:latin'), 'typeface')).toBe('Georgia');
  });
});

describe('PPTX exporter — oversized text-run unit spills onto a continuation slide', () => {
  // Many short, consecutive <p>s with no images/cards between them all merge
  // into ONE text-run "unit" (classifyContentUnits' own pushText grouping) —
  // exactly what a scrollable HTML container's paragraphs would produce. Before
  // the slide.js fix, a unit like this that didn't fit a slide's remaining
  // space rendered in full at cursor===0 and silently overflowed past the
  // slide's own bottom edge instead of continuing onto a fresh slide.
  const manyParagraphs = Array.from({ length: 60 }, (_, i) => `<p>Line number ${i} of body text, long enough to add real height.</p>`).join('');

  it('splits the oversized unit across as many slides as its content needs, keeping every paragraph', async () => {
    const result = await convert(manyParagraphs, { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    const slideCount = [...zip.entries()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
    expect(slideCount).toBeGreaterThan(1);

    const allText = Array.from({ length: slideCount }, (_, i) => slideTexts(xml[`ppt/slides/slide${i + 1}.xml`]).join(' ')).join(' ');
    for (let i = 0; i < 60; i++) {
      expect(allText).toContain(`Line number ${i} of`);
    }
  });

  it('a single slide holds every paragraph when they collectively fit', async () => {
    const result = await convert('<p>Short one</p><p>Short two</p>', { format: 'pptx' });
    const { zip, xml } = readPptxParts(result.buffer);
    const slideCount = [...zip.entries()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
    expect(slideCount).toBe(1);
    expect(slideTexts(xml['ppt/slides/slide1.xml'])).toEqual(expect.arrayContaining(['Short one', 'Short two']));
  });
});
