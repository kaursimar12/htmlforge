/**
 * xl/styles.xml — cell formatting. Excel requires non-empty fonts/fills
 * (including the 2 reserved fills)/borders/cellStyleXfs/cellXfs/cellStyles
 * even for a file with no real styling at all; `CellStyleManager` builds
 * that minimal skeleton and layers a deduplicating composite registry on
 * top (`registerCellStyle`) — one font/fill/border/numFmt sub-table each,
 * combined into a `cellXfs` entry whose index is the `s="N"` a cell uses.
 *
 * Index 0 (no bold, no fill, no border, General format, no alignment) is
 * pre-registered so an untouched cell can just omit `s` entirely and still
 * mean exactly this.
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';
import { fontStyleKey } from './font-style-key.js';
import { resolveFontFamilyName } from '../../fonts/resolver.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

// CSS border-style values that have a direct ST_BorderStyle equivalent —
// mirrors the OOXML_BORDER_STYLE mapping already used for DOCX's WordprocessingML
// borders, but SpreadsheetML's own vocabulary distinguishes thin/medium/thick
// by width rather than having a single generic "single" value.
function excelBorderStyle(border) {
  if (border.style === 'dashed') return 'dashed';
  if (border.style === 'dotted') return 'dotted';
  if (border.style === 'double') return 'double';
  // 'single', or anything unrecognized: bucket by width.
  if (border.sizeEighthPt <= 8) return 'thin';
  if (border.sizeEighthPt <= 16) return 'medium';
  return 'thick';
}

// Custom number formats live at id >= 164 — 0-163 are reserved for Excel's
// own builtins (0 = "General", the only builtin this project ever assigns).
const FIRST_CUSTOM_NUM_FMT_ID = 164;


export class CellStyleManager {
  // The document's own `body{font-family|font-size}` (dom-to-idm.js's
  // metadata.fontFamily/fontSizePt) — every plain, otherwise-unstyled cell
  // (fontId 0, the pre-registered index-0 entry above) should render at
  // this, not this project's own hardcoded Calibri-11pt fallback. A cell
  // with its own explicit fontSizePt (registerCellStyle's own param) still
  // overrides this, same "explicit beats inferred" precedence used
  // throughout this exporter — see #registerFont/toXml below. There is no
  // per-cell fontFamily override anywhere in this exporter yet, so the
  // document default is applied to every font entry uniformly.
  #defaultFontFamily;
  #defaultFontSizePt;
  #fonts = [{ key: fontStyleKey({}) }];
  #fills = [{ key: 'none' }, { key: 'gray125' }]; // indices 0/1 are schema-reserved
  #borders = [{ key: JSON.stringify({}), sides: {} }];
  #numFmts = []; // { id, code }
  #numFmtIndex = new Map(); // code -> id
  #cellXfs = [{ key: JSON.stringify({ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, align: undefined }), fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, align: undefined }];
  #xfIndex = new Map([[this.#cellXfs[0].key, 0]]);

  constructor({ fontFamily, fontSizePt } = {}) {
    // Same generic-keyword mapping DOCX/PPTX apply (resolveFontFamilyName) —
    // XLSX had been writing a bare CSS value like "sans-serif"/"monospace"
    // straight into `<name val="...">` unchanged, which isn't a real
    // installable font name at all, unlike an arbitrary custom family
    // (e.g. "Poppins"), which still passes through untouched.
    this.#defaultFontFamily = resolveFontFamilyName(fontFamily);
    this.#defaultFontSizePt = fontSizePt;
  }

  #registerFont({ bold, italic, underline, strikethrough, color, fontSizePt }) {
    const key = fontStyleKey({ bold, italic, underline, strikethrough, color, fontSizePt });
    const existing = this.#fonts.findIndex((f) => f.key === key);
    if (existing !== -1) return existing;
    this.#fonts.push({ key, bold: !!bold, italic: !!italic, underline: !!underline, strikethrough: !!strikethrough, color, fontSizePt });
    return this.#fonts.length - 1;
  }

  #registerFill(backgroundColor) {
    if (!backgroundColor) return 0;
    const key = backgroundColor;
    const existing = this.#fills.findIndex((f) => f.key === key);
    if (existing !== -1) return existing;
    this.#fills.push({ key, color: backgroundColor });
    return this.#fills.length - 1;
  }

  #registerBorder(sides) {
    const hasAny = sides.top || sides.right || sides.bottom || sides.left;
    if (!hasAny) return 0;
    const key = JSON.stringify(sides);
    const existing = this.#borders.findIndex((b) => b.key === key);
    if (existing !== -1) return existing;
    this.#borders.push({ key, sides });
    return this.#borders.length - 1;
  }

  #registerNumFmt(code) {
    if (!code) return 0;
    const existing = this.#numFmtIndex.get(code);
    if (existing !== undefined) return existing;
    const id = FIRST_CUSTOM_NUM_FMT_ID + this.#numFmts.length;
    this.#numFmts.push({ id, code });
    this.#numFmtIndex.set(code, id);
    return id;
  }

  /**
   * Registers (deduping) a composite cell style and returns its `s="N"`
   * cellXfs index. Every field is optional; an empty/all-falsy call
   * resolves back to the pre-registered default index 0.
   */
  registerCellStyle({ bold, italic, underline, strikethrough, color, fontSizePt, backgroundColor, borderTop, borderRight, borderBottom, borderLeft, numFmtCode, align, verticalAlign, readingOrder, textRotation, wrapText } = {}) {
    const fontId = this.#registerFont({ bold, italic, underline, strikethrough, color, fontSizePt });
    const fillId = this.#registerFill(backgroundColor);
    const borderId = this.#registerBorder({ top: borderTop, right: borderRight, bottom: borderBottom, left: borderLeft });
    const numFmtId = this.#registerNumFmt(numFmtCode);

    // `wrapText || undefined` (not `!!wrapText`) so a falsy/absent value drops
    // out of the JSON key entirely, same as every other optional field here —
    // keeps the untouched-cell key identical to the pre-registered index 0
    // entry below, which predates this field and has no `wrapText` of its own.
    const key = JSON.stringify({ fontId, fillId, borderId, numFmtId, align, verticalAlign, readingOrder, textRotation, wrapText: wrapText || undefined });
    const existing = this.#xfIndex.get(key);
    if (existing !== undefined) return existing;
    const index = this.#cellXfs.length;
    this.#cellXfs.push({ key, fontId, fillId, borderId, numFmtId, align, verticalAlign, readingOrder, textRotation, wrapText: !!wrapText });
    this.#xfIndex.set(key, index);
    return index;
  }

  toXml() {
    const b = new XmlBuilder();
    b.el('styleSheet', { xmlns: NS_MAIN });

    if (this.#numFmts.length > 0) {
      b.el('numFmts', { count: this.#numFmts.length });
      for (const { id, code } of this.#numFmts) b.self('numFmt', { numFmtId: id, formatCode: code });
      b.up();
    }

    b.el('fonts', { count: this.#fonts.length });
    for (const font of this.#fonts) {
      b.el('font');
      if (font.bold) b.self('b');
      if (font.italic) b.self('i');
      if (font.underline) b.self('u');
      if (font.strikethrough) b.self('strike');
      if (font.color) b.self('color', { rgb: `FF${font.color}` });
      b.self('sz', { val: font.fontSizePt || this.#defaultFontSizePt || 11 });
      b.self('name', { val: this.#defaultFontFamily || 'Calibri' });
      b.up();
    }
    b.up();

    b.el('fills', { count: this.#fills.length });
    b.el('fill').self('patternFill', { patternType: 'none' }).up();
    b.el('fill').self('patternFill', { patternType: 'gray125' }).up();
    for (const fill of this.#fills.slice(2)) {
      b.el('fill');
      b.el('patternFill', { patternType: 'solid' });
      b.self('fgColor', { rgb: `FF${fill.color}` });
      b.up();
      b.up();
    }
    b.up();

    b.el('borders', { count: this.#borders.length });
    for (const border of this.#borders) {
      b.el('border');
      for (const side of ['left', 'right', 'top', 'bottom']) {
        const s = border.sides?.[side];
        if (s) {
          b.el(side, { style: excelBorderStyle(s) });
          b.self('color', { rgb: `FF${s.color}` });
          b.up();
        } else {
          b.self(side);
        }
      }
      b.self('diagonal');
      b.up();
    }
    b.up();

    b.el('cellStyleXfs', { count: 1 });
    b.self('xf', { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
    b.up();

    b.el('cellXfs', { count: this.#cellXfs.length });
    for (const xf of this.#cellXfs) {
      const attrs = { numFmtId: xf.numFmtId, fontId: xf.fontId, fillId: xf.fillId, borderId: xf.borderId, xfId: 0 };
      if (xf.fontId) attrs.applyFont = 1;
      if (xf.fillId) attrs.applyFill = 1;
      if (xf.borderId) attrs.applyBorder = 1;
      const hasAlignment = xf.align || xf.verticalAlign || xf.readingOrder || xf.textRotation || xf.wrapText;
      if (hasAlignment) attrs.applyAlignment = 1;
      if (hasAlignment) {
        b.el('xf', attrs);
        const alignAttrs = {};
        if (xf.align) alignAttrs.horizontal = xf.align;
        if (xf.verticalAlign) alignAttrs.vertical = xf.verticalAlign;
        // `direction:rtl` — 2 = right-to-left reading order (1 = LTR, 0 = context-dependent, Excel's default).
        if (xf.readingOrder) alignAttrs.readingOrder = xf.readingOrder;
        // Already-computed ST_CellAlignmentTextRotation value — see
        // worksheet.js's cellTextRotation for how vertical-rl/rotate(Ndeg) map to it.
        if (xf.textRotation !== undefined) alignAttrs.textRotation = xf.textRotation;
        // A block's text carries literal `\n`s (from <br>) that only render as
        // real line breaks in Excel when the cell also has wrapText set.
        if (xf.wrapText) alignAttrs.wrapText = 1;
        b.self('alignment', alignAttrs);
        b.up();
      } else {
        b.self('xf', attrs);
      }
    }
    b.up();

    b.el('cellStyles', { count: 1 });
    b.self('cellStyle', { name: 'Normal', xfId: 0, builtinId: 0 });
    b.up();

    return xmlDoc(b.end());
  }
}
