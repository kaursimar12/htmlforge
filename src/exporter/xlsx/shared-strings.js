/**
 * xl/sharedStrings.xml — the workbook-wide deduplicated string table. One
 * instance is shared across every sheet in an export (not one per sheet),
 * so the same text on two different sheets reuses a single entry. Supports
 * both plain strings and rich-text entries (multiple differently-styled
 * runs — see rich-text.js) in the same table.
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export class SharedStringsManager {
  #strings = []; // string | { rich: Run[] }
  #index = new Map(); // dedup key -> its index in #strings
  #refCount = 0; // total number of t="s" cell references across the workbook

  /** Registers (deduping) a plain string and returns its shared-string index. */
  add(value) {
    return this.#register(`s:${value}`, value);
  }

  /**
   * Registers (deduping) a rich-text entry — an array of `{text, bold?,
   * italic?, underline?, strikethrough?, color?, fontSizePt?}` runs — and
   * returns its shared-string index.
   */
  addRich(runs) {
    return this.#register(`r:${JSON.stringify(runs)}`, { rich: runs });
  }

  #register(key, entry) {
    this.#refCount++;
    const existing = this.#index.get(key);
    if (existing !== undefined) return existing;
    const index = this.#strings.length;
    this.#strings.push(entry);
    this.#index.set(key, index);
    return index;
  }

  toXml() {
    const b = new XmlBuilder();
    b.el('sst', { xmlns: NS_MAIN, count: this.#refCount, uniqueCount: this.#strings.length });
    for (const entry of this.#strings) {
      b.el('si');
      if (typeof entry === 'string') {
        // xml:space="preserve" so Excel doesn't collapse/trim significant
        // leading/trailing whitespace out of the displayed text.
        b.el('t', { 'xml:space': 'preserve' }).text(entry).up();
      } else {
        for (const run of entry.rich) renderRun(b, run);
      }
      b.up();
    }
    return xmlDoc(b.end());
  }
}

// A run's rPr uses the same font sub-elements as xl/styles.xml's <font> —
// except font-family is `rFont` here (CT_RPrElt), not `name` (CT_Font);
// that's a genuine, deliberate OOXML naming difference between the two
// contexts, not an inconsistency in this codebase. Only emitted when the
// run actually carries an explicit property — a plain run (the common case
// within an otherwise-styled cell) has no <rPr> at all and just inherits
// the cell's own default font.
function renderRun(b, run) {
  b.el('r');
  const hasStyle = run.bold || run.italic || run.underline || run.strikethrough || run.color || run.fontSizePt;
  if (hasStyle) {
    b.el('rPr');
    if (run.bold) b.self('b');
    if (run.italic) b.self('i');
    if (run.underline) b.self('u');
    if (run.strikethrough) b.self('strike');
    if (run.fontSizePt) b.self('sz', { val: run.fontSizePt });
    if (run.color) b.self('color', { rgb: `FF${run.color}` });
    b.up();
  }
  b.el('t', { 'xml:space': 'preserve' }).text(run.text).up();
  b.up();
}
