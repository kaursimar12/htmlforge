/**
 * xl/drawings/drawing1.xml (X8 — images) — one <xdr:oneCellAnchor> per
 * embedded image, anchored at the {row, col} document-sheet.js reserved for
 * it. Structurally the SpreadsheetML sibling of DOCX's own inline picture
 * run (docx/document.js's renderImageRun) — same <a:blip>/<a:stretch>/
 * <a:xfrm> picture shape, just anchored to a worksheet cell instead of
 * inlined into a paragraph run.
 */
import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';

const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

// 914400 EMU/in ÷ 96 CSS px/in — same reference docx/document.js's identical
// EMU_PER_PX uses for its own <a:ext>/<wp:extent> sizing.
const EMU_PER_PX = 9525;

/** `anchors`: `[{ row, col, widthPx, heightPx, relId }]`, one per embedded image. */
export function drawingXml(anchors) {
  const b = new XmlBuilder();
  b.el('xdr:wsDr', { 'xmlns:xdr': NS_XDR, 'xmlns:a': NS_A, 'xmlns:r': NS_R });

  anchors.forEach(({ row, col, widthPx, heightPx, relId }, i) => {
    const cx = Math.round(widthPx * EMU_PER_PX);
    const cy = Math.round(heightPx * EMU_PER_PX);
    const name = `Image${i + 1}`;

    b.el('xdr:oneCellAnchor');
    b.el('xdr:from');
    b.el('xdr:col').text(String(col)).up();
    b.el('xdr:colOff').text('0').up();
    b.el('xdr:row').text(String(row)).up();
    b.el('xdr:rowOff').text('0').up();
    b.up(); // xdr:from
    b.self('xdr:ext', { cx, cy });

    b.el('xdr:pic');
    b.el('xdr:nvPicPr');
    b.self('xdr:cNvPr', { id: i + 2, name });
    b.self('xdr:cNvPicPr');
    b.up(); // xdr:nvPicPr

    b.el('xdr:blipFill');
    b.self('a:blip', { 'r:embed': relId });
    b.el('a:stretch');
    b.self('a:fillRect');
    b.up(); // a:stretch
    b.up(); // xdr:blipFill

    b.el('xdr:spPr');
    b.el('a:xfrm');
    b.self('a:off', { x: 0, y: 0 });
    b.self('a:ext', { cx, cy });
    b.up(); // a:xfrm
    b.el('a:prstGeom', { prst: 'rect' });
    b.self('a:avLst');
    b.up(); // a:prstGeom
    b.up(); // xdr:spPr
    b.up(); // xdr:pic

    b.self('xdr:clientData');
    b.up(); // xdr:oneCellAnchor
  });

  return xmlDoc(b.end());
}

/** `anchors`: same shape as drawingXml's, plus each anchor's own `filename`. */
export function drawingRelsXml(anchors) {
  const b = new XmlBuilder();
  b.el('Relationships', { xmlns: NS_PACKAGE_RELS });
  for (const { relId, filename } of anchors) {
    b.self('Relationship', { Id: relId, Type: `${NS_R}/image`, Target: `../media/${filename}` });
  }
  return xmlDoc(b.end());
}
