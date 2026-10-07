/**
 * xl/workbook.xml (sheet list) + xl/_rels/workbook.xml.rels + the one
 * sheet's name. The whole document renders as a single worksheet (see
 * document-sheet.js) rather than one sheet per `<table>`.
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';
import { RelationshipManager, REL_TYPES } from '../ooxml/package.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const ILLEGAL_SHEET_NAME_CHARS = /[\\/?*[\]:]/g;
const MAX_SHEET_NAME_LENGTH = 31;

export function sanitizeSheetName(name, fallbackIndex) {
  const cleaned = name.replace(ILLEGAL_SHEET_NAME_CHARS, ' ').replace(/\s+/g, ' ').trim();
  return cleaned || `Sheet${fallbackIndex}`;
}

// The whole document renders as one sheet — name it from convert()'s
// sheetName option, then the document's own <title>, then (only when the
// document is a single table with no other content) that table's own
// <caption>, falling back to "Sheet1". Sanitized to Excel's legal character
// set and truncated to its 31-char limit.
export function resolveDocumentSheetName(document, tables, options) {
  const raw = options?.sheetName || document?.metadata?.title || (tables.length === 1 ? tables[0].caption : undefined) || 'Sheet1';
  return sanitizeSheetName(raw.slice(0, MAX_SHEET_NAME_LENGTH), 1);
}

/** Registers styles.xml/sharedStrings.xml + one worksheet rel per sheet. */
export function buildWorkbookRels(sheetCount) {
  const rels = new RelationshipManager();
  rels.add(REL_TYPES.styles, 'styles.xml');
  rels.add(REL_TYPES.sharedStrings, 'sharedStrings.xml');
  const sheetRelIds = [];
  for (let i = 0; i < sheetCount; i++) {
    sheetRelIds.push(rels.add(REL_TYPES.worksheet, `worksheets/sheet${i + 1}.xml`));
  }
  return { xml: rels.toXml(), sheetRelIds };
}

/** `sheets` is `[{ name, rId }]`, in the same order as their sheetN.xml files. */
export function workbookXml(sheets) {
  const b = new XmlBuilder();
  b.el('workbook', { xmlns: NS_MAIN, 'xmlns:r': NS_R });
  b.el('sheets');
  sheets.forEach((sheet, i) => b.self('sheet', { name: sheet.name, sheetId: i + 1, 'r:id': sheet.rId }));
  b.up();
  b.up();
  return xmlDoc(b.end());
}
