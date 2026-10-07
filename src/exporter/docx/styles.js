/**
 * word/styles.xml — built up incrementally as milestones land:
 * W1 Normal · W3 Heading1-6 · W4 Quote · W5 ListBullet/ListNumber ·
 * W6 TableGrid · W7 Hyperlink
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';
import { resolveFontFamilyName } from '../../fonts/resolver.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Half-points (Word's unit for w:sz), decreasing per level.
const HEADING_SIZES = [36, 32, 28, 26, 24, 22];

// ListBullet/ListNumber differ only in their id/display name — both are a
// plain paragraph style based on Normal with spacing zeroed out (see the
// call site's own comment on why).
function addZeroSpacingListStyle(b, styleId, name) {
  b.el('w:style', { 'w:type': 'paragraph', 'w:styleId': styleId });
  b.self('w:name', { 'w:val': name });
  b.self('w:basedOn', { 'w:val': 'Normal' });
  b.el('w:pPr');
  b.self('w:spacing', { 'w:before': 0, 'w:after': 0 });
  b.up();
  b.up();
}

export function stylesXml(metadata = {}) {
  const b = new XmlBuilder();
  b.el('w:styles', { 'xmlns:w': W_NS });

  // Without an explicit "space after," every paragraph in a document built
  // entirely from scratch (no inherited Normal.dotm defaults) renders with
  // zero gap between them — every real HTML→DOCX converter (and Word's own
  // default template) gives paragraphs breathing room by default; 160 twips
  // (8pt) matches Word's own common default and roughly matches a plain
  // <p>'s browser UA-stylesheet margin. An explicit CSS margin on a specific
  // element still overrides this via direct formatting (see renderSpacing).
  b.el('w:style', { 'w:type': 'paragraph', 'w:default': true, 'w:styleId': 'Normal' });
  b.self('w:name', { 'w:val': 'Normal' });
  b.el('w:pPr');
  b.self('w:spacing', { 'w:after': 160 });
  b.up();
  // The document's own `body{font-family|font-size}` (dom-to-idm.js's
  // metadata.fontFamily/fontSizePt) — every plain, otherwise-unstyled run
  // (including ListBullet/ListNumber/Quote/table-cell text, all basedOn
  // Normal) should render at this, not Word's own built-in default (11pt
  // Calibri). An element with its own explicit fontFamily/fontSizePt (a
  // heading's own w:sz above, a styled span's own w:rPr — see
  // document.js's renderRPr) still overrides this via direct formatting,
  // same "explicit beats inferred" precedence used throughout this
  // exporter. Omitted entirely when the source declared neither, leaving
  // Word's built-in default exactly as before this fix.
  if (metadata.fontFamily || metadata.fontSizePt) {
    b.el('w:rPr');
    if (metadata.fontFamily) {
      const resolvedFontFamily = resolveFontFamilyName(metadata.fontFamily);
      b.self('w:rFonts', { 'w:ascii': resolvedFontFamily, 'w:hAnsi': resolvedFontFamily });
    }
    if (metadata.fontSizePt) b.self('w:sz', { 'w:val': Math.round(metadata.fontSizePt * 2) });
    b.up();
  }
  b.up();

  HEADING_SIZES.forEach((sz, i) => {
    const level = i + 1;
    b.el('w:style', { 'w:type': 'paragraph', 'w:styleId': `Heading${level}` });
    b.self('w:name', { 'w:val': `heading ${level}` });
    b.self('w:basedOn', { 'w:val': 'Normal' });
    b.self('w:next', { 'w:val': 'Normal' });
    b.el('w:pPr');
    b.self('w:outlineLvl', { 'w:val': i });
    // A heading needs more separation from the content above it than a
    // plain paragraph does — matching a browser's default h1-h6 margins,
    // which are consistently larger than a <p>'s.
    b.self('w:spacing', { 'w:before': 240, 'w:after': 120 });
    b.up();
    b.el('w:rPr');
    b.self('w:b');
    b.self('w:sz', { 'w:val': sz });
    b.up();
    b.up();
  });

  b.el('w:style', { 'w:type': 'paragraph', 'w:styleId': 'Quote' });
  b.self('w:name', { 'w:val': 'Quote' });
  b.self('w:basedOn', { 'w:val': 'Normal' });
  b.el('w:pPr');
  b.self('w:ind', { 'w:left': 720, 'w:right': 720 });
  b.up();
  b.el('w:rPr');
  b.self('w:i');
  b.up();
  b.up();

  // A browser renders adjacent <li> items with no gap between them (only
  // the list as a whole gets top/bottom margin) — overriding Normal's
  // default spacing here keeps that same tight, correct look instead of
  // inheriting an 8pt gap between every bullet/number.
  addZeroSpacingListStyle(b, 'ListBullet', 'List Bullet');
  addZeroSpacingListStyle(b, 'ListNumber', 'List Number');

  b.el('w:style', { 'w:type': 'character', 'w:styleId': 'Hyperlink' });
  b.self('w:name', { 'w:val': 'Hyperlink' });
  b.el('w:rPr');
  b.self('w:color', { 'w:val': '0563C1' });
  b.self('w:u', { 'w:val': 'single' });
  b.up();
  b.up();

  b.el('w:style', { 'w:type': 'table', 'w:styleId': 'TableGrid' });
  b.self('w:name', { 'w:val': 'Table Grid' });
  b.self('w:basedOn', { 'w:val': 'TableNormal' });
  b.el('w:tblPr');
  b.el('w:tblBorders');
  for (const edge of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
    b.self(`w:${edge}`, { 'w:val': 'single', 'w:sz': 4, 'w:space': 0, 'w:color': 'auto' });
  }
  b.up();
  b.up();
  b.up();

  return xmlDoc(b.end());
}
