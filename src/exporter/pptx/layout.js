/**
 * PPTX slide geometry — one fixed EMU canvas per presentation (no per-slide
 * layout engine), plus the static theme/master/layout parts every .pptx
 * needs regardless of content. PowerPoint (unlike Word) shows a "repair"
 * prompt for a slide deck missing a theme or a master/layout pair — pptxgenjs
 * always ships a full theme1.xml for exactly this reason, so this does too,
 * kept minimal (default Office-like scheme) rather than empty.
 */
import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';
import { NS, RelationshipManager } from '../ooxml/package.js';

export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const NS_R = NS.officeRels;

export const PPTX_REL_TYPES = {
  slideMaster: `${NS.officeRels}/slideMaster`,
  slideLayout: `${NS.officeRels}/slideLayout`,
  theme: `${NS.officeRels}/theme`,
  slide: `${NS.officeRels}/slide`,
  image: `${NS.officeRels}/image`,
  hyperlink: `${NS.officeRels}/hyperlink`,
};

// EMUs (914400 per inch) — matches the constant already used in docx/document.js.
export const EMU_PER_IN = 914400;
export const SLIDE_SIZE = {
  widescreen: { cx: 12192000, cy: 6858000 }, // 16:9, PowerPoint's modern default
  standard: { cx: 9144000, cy: 6858000 }, // 4:3
};

export function resolveSlideSize(options) {
  return SLIDE_SIZE[options?.slideSize] ?? SLIDE_SIZE.widescreen;
}

// Fixed proportions of the slide canvas — not a real layout engine (none of
// the OOXML exporters in this project have one), but enough to place a title,
// a body text placeholder, and any media shapes without overlap.
export function computeSlideGeometry(slideSize) {
  const margin = Math.round(EMU_PER_IN * 0.5);
  const contentWidth = slideSize.cx - margin * 2;
  const titleY = Math.round(EMU_PER_IN * 0.3);
  const titleHeight = Math.round(EMU_PER_IN * 1.25);
  const bodyY = titleY + titleHeight + Math.round(EMU_PER_IN * 0.15);
  const bottomY = slideSize.cy - margin;
  return { margin, contentWidth, titleY, titleHeight, bodyY, bottomY };
}

export function presentationXml(masterRId, slideRIds, slideSize) {
  const b = new XmlBuilder();
  b.el('p:presentation', { 'xmlns:a': NS_A, 'xmlns:p': NS_P, 'xmlns:r': NS_R });
  b.el('p:sldMasterIdLst');
  b.self('p:sldMasterId', { id: 2147483648, 'r:id': masterRId });
  b.up();
  b.el('p:sldIdLst');
  slideRIds.forEach((rId, i) => b.self('p:sldId', { id: 256 + i, 'r:id': rId }));
  b.up();
  b.self('p:sldSz', { cx: slideSize.cx, cy: slideSize.cy, type: slideSize === SLIDE_SIZE.standard ? 'screen4x3' : 'screen16x9' });
  // Notes page size is independent of the slide canvas — a fixed portrait
  // page (6858000x9144000 EMU, ~7.5x10in) matches PowerPoint's own default.
  b.self('p:notesSz', { cx: 6858000, cy: 9144000 });
  b.up();
  return xmlDoc(b.end());
}

/** presentation.xml.rels — master first (fixed rId), then one entry per slide, in order. */
export function buildPresentationRels(slideCount) {
  const rels = new RelationshipManager();
  // RelationshipManager auto-allocates rId1, rId2, ... sequentially; the
  // master is registered first purely so its id is stable/predictable, but
  // presentationXml above references it by the literal id returned here.
  const masterRId = rels.add(PPTX_REL_TYPES.slideMaster, 'slideMasters/slideMaster1.xml');
  const slideRIds = [];
  for (let i = 0; i < slideCount; i++) {
    slideRIds.push(rels.add(PPTX_REL_TYPES.slide, `slides/slide${i + 1}.xml`));
  }
  return { xml: rels.toXml(), masterRId, slideRIds };
}

// A minimal but complete Office-default-like theme — required for
// PowerPoint/Impress to open the deck without a repair prompt. Colors are
// purely cosmetic, but the font is not: every run that doesn't declare its
// own font-family (the overwhelming majority — see shapes.js's renderRPr,
// which only ever emits `a:latin` for a code run or an explicit CSS
// font-family) inherits whichever typeface is named here, so this is the
// deck-wide default typeface a viewer actually sees. Defaults to Calibri —
// PowerPoint's own default — but a source document's own `body{font-family:
// ...}` (see dom-to-idm.js's domToIdm, metadata.fontFamily) should win over
// that generic default, the same "explicit beats inferred" precedence used
// everywhere else in this project.
export function themeXml(defaultFontFamily = 'Calibri') {
  const b = new XmlBuilder();
  b.el('a:theme', { 'xmlns:a': NS_A, name: 'html-exporter Theme' });
  b.el('a:themeElements');

  b.el('a:clrScheme', { name: 'html-exporter' });
  const colors = [
    ['a:dk1', 'sysClr', { val: 'windowText', lastClr: '000000' }],
    ['a:lt1', 'sysClr', { val: 'window', lastClr: 'FFFFFF' }],
    ['a:dk2', 'srgbClr', { val: '44546A' }],
    ['a:lt2', 'srgbClr', { val: 'E7E6E6' }],
    ['a:accent1', 'srgbClr', { val: '4472C4' }],
    ['a:accent2', 'srgbClr', { val: 'ED7D31' }],
    ['a:accent3', 'srgbClr', { val: 'A5A5A5' }],
    ['a:accent4', 'srgbClr', { val: 'FFC000' }],
    ['a:accent5', 'srgbClr', { val: '5B9BD5' }],
    ['a:accent6', 'srgbClr', { val: '70AD47' }],
    ['a:hlink', 'srgbClr', { val: '0563C1' }],
    ['a:folHlink', 'srgbClr', { val: '954F72' }],
  ];
  for (const [tag, childTag, attrs] of colors) {
    b.el(tag);
    b.self(childTag, attrs);
    b.up();
  }
  b.up(); // a:clrScheme

  b.el('a:fontScheme', { name: 'html-exporter' });
  for (const scheme of ['a:majorFont', 'a:minorFont']) {
    b.el(scheme);
    b.self('a:latin', { typeface: defaultFontFamily });
    b.self('a:ea', { typeface: '' });
    b.self('a:cs', { typeface: '' });
    b.up();
  }
  b.up(); // a:fontScheme

  b.el('a:fmtScheme', { name: 'html-exporter' });
  b.el('a:fillStyleLst');
  b.el('a:solidFill').self('a:schemeClr', { val: 'phClr' }).up();
  b.el('a:solidFill').self('a:schemeClr', { val: 'phClr' }).up();
  b.el('a:solidFill').self('a:schemeClr', { val: 'phClr' }).up();
  b.up();
  b.el('a:lnStyleLst');
  for (let i = 0; i < 3; i++) {
    b.el('a:ln', { w: 6350 + i * 6350 });
    b.el('a:solidFill').self('a:schemeClr', { val: 'phClr' }).up();
    b.up();
  }
  b.up();
  b.el('a:effectStyleLst');
  for (let i = 0; i < 3; i++) { b.el('a:effectStyle'); b.el('a:effectLst').up(); b.up(); }
  b.up();
  b.el('a:bgFillStyleLst');
  b.el('a:solidFill').self('a:schemeClr', { val: 'phClr' }).up();
  b.el('a:solidFill').self('a:schemeClr', { val: 'phClr' }).up();
  b.el('a:solidFill').self('a:schemeClr', { val: 'phClr' }).up();
  b.up();
  b.up(); // a:fmtScheme

  b.up(); // a:themeElements
  return xmlDoc(b.end());
}

function placeholderShape(b, { shapeId, phType, phIdx, name, x, y, cx, cy }) {
  b.el('p:sp');
  b.el('p:nvSpPr');
  b.self('p:cNvPr', { id: shapeId, name });
  b.self('p:cNvSpPr');
  b.el('p:nvPr');
  b.self('p:ph', { type: phType, idx: phIdx });
  b.up();
  b.up(); // p:nvSpPr
  b.el('p:spPr');
  if (x !== undefined) {
    b.el('a:xfrm');
    b.self('a:off', { x, y });
    b.self('a:ext', { cx, cy });
    b.up();
  }
  b.up();
  b.el('p:txBody');
  b.self('a:bodyPr');
  b.self('a:lstStyle');
  b.el('a:p').up();
  b.up();
  b.up(); // p:sp
}

export function slideMasterXml(slideSize, layoutRId) {
  const geo = computeSlideGeometry(slideSize);
  const b = new XmlBuilder();
  b.el('p:sldMaster', { 'xmlns:a': NS_A, 'xmlns:p': NS_P, 'xmlns:r': NS_R });
  b.el('p:cSld');
  b.el('p:spTree');
  b.el('p:nvGrpSpPr');
  b.self('p:cNvPr', { id: 1, name: '' });
  b.self('p:cNvGrpSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:grpSpPr').up();
  placeholderShape(b, { shapeId: 2, phType: 'title', name: 'Title Placeholder', x: geo.margin, y: geo.titleY, cx: geo.contentWidth, cy: geo.titleHeight });
  placeholderShape(b, { shapeId: 3, phType: 'body', phIdx: 1, name: 'Body Placeholder', x: geo.margin, y: geo.bodyY, cx: geo.contentWidth, cy: geo.bottomY - geo.bodyY });
  b.up(); // p:spTree
  b.up(); // p:cSld
  b.self('p:clrMap', {
    bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2',
    accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
    hlink: 'hlink', folHlink: 'folHlink',
  });
  b.el('p:sldLayoutIdLst');
  b.self('p:sldLayoutId', { id: 2147483649, 'r:id': layoutRId });
  b.up();
  return xmlDoc(b.end());
}

export function buildSlideMasterRels() {
  const rels = new RelationshipManager();
  const layoutRId = rels.add(PPTX_REL_TYPES.slideLayout, '../slideLayouts/slideLayout1.xml');
  rels.add(PPTX_REL_TYPES.theme, '../theme/theme1.xml');
  return { xml: rels.toXml(), layoutRId };
}

export function slideLayoutXml(slideSize) {
  const geo = computeSlideGeometry(slideSize);
  const b = new XmlBuilder();
  b.el('p:sldLayout', { 'xmlns:a': NS_A, 'xmlns:p': NS_P, 'xmlns:r': NS_R, type: 'title', preserve: 1 });
  b.el('p:cSld', { name: 'Title and Body' });
  b.el('p:spTree');
  b.el('p:nvGrpSpPr');
  b.self('p:cNvPr', { id: 1, name: '' });
  b.self('p:cNvGrpSpPr');
  b.self('p:nvPr');
  b.up();
  b.el('p:grpSpPr').up();
  placeholderShape(b, { shapeId: 2, phType: 'title', name: 'Title Placeholder', x: geo.margin, y: geo.titleY, cx: geo.contentWidth, cy: geo.titleHeight });
  placeholderShape(b, { shapeId: 3, phType: 'body', phIdx: 1, name: 'Body Placeholder', x: geo.margin, y: geo.bodyY, cx: geo.contentWidth, cy: geo.bottomY - geo.bodyY });
  b.up(); // p:spTree
  b.up(); // p:cSld
  return xmlDoc(b.end());
}

export function slideLayoutRelsXml() {
  const rels = new RelationshipManager();
  rels.add(PPTX_REL_TYPES.slideMaster, '../slideMasters/slideMaster1.xml');
  return rels.toXml();
}
