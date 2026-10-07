/**
 * word/numbering.xml — list numbering definitions (W5).
 *
 * Two shared abstractNum definitions (one bullet, one decimal-family), each
 * with all 9 levels pre-defined. Every distinct <ul>/<ol> in the document
 * gets its own <w:num> entry pointing at the matching abstractNumId — this
 * is what lets Word treat two separate lists as independently numbered.
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const BULLET_ABSTRACT_ID = 0;
const DECIMAL_ABSTRACT_ID = 1;

const BULLET_CHARS = ['•', '◦', '▪'];
const ORDERED_FORMATS = ['decimal', 'lowerLetter', 'lowerRoman'];

export class NumberingManager {
  #nums = [];
  #counter = 1;

  /** Registers a new <ul>/<ol> and returns the numId it should use. */
  registerList(ordered, start) {
    const numId = this.#counter++;
    this.#nums.push({ numId, abstractNumId: ordered ? DECIMAL_ABSTRACT_ID : BULLET_ABSTRACT_ID, start });
    return numId;
  }

  toXml() {
    const b = new XmlBuilder();
    b.el('w:numbering', { 'xmlns:w': W_NS });
    b.raw(abstractNumXml(BULLET_ABSTRACT_ID, 'bullet', BULLET_CHARS));
    b.raw(abstractNumXml(DECIMAL_ABSTRACT_ID, 'ordered', ORDERED_FORMATS));
    for (const n of this.#nums) {
      b.el('w:num', { 'w:numId': n.numId });
      b.self('w:abstractNumId', { 'w:val': n.abstractNumId });
      if (n.start !== undefined && n.start !== 1) {
        b.el('w:lvlOverride', { 'w:ilvl': 0 });
        b.self('w:startOverride', { 'w:val': n.start });
        b.up();
      }
      b.up();
    }
    return xmlDoc(b.end());
  }
}

function abstractNumXml(id, kind, cycle) {
  const b = new XmlBuilder();
  b.el('w:abstractNum', { 'w:abstractNumId': id });
  for (let lvl = 0; lvl < 9; lvl++) {
    b.el('w:lvl', { 'w:ilvl': lvl });
    b.self('w:start', { 'w:val': 1 });
    if (kind === 'bullet') {
      b.self('w:numFmt', { 'w:val': 'bullet' });
      b.self('w:lvlText', { 'w:val': cycle[lvl % cycle.length] });
    } else {
      b.self('w:numFmt', { 'w:val': cycle[lvl % cycle.length] });
      // Each level's own counter (%N references itself, not a parent level) —
      // matches how a real multi-level list restarts per level.
      b.self('w:lvlText', { 'w:val': `%${lvl + 1}.` });
    }
    b.self('w:lvlJc', { 'w:val': 'left' });
    b.el('w:pPr');
    b.self('w:ind', { 'w:left': 720 * (lvl + 1), 'w:hanging': 360 });
    b.up();
    b.up();
  }
  b.up();
  return b.end();
}
