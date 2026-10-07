/**
 * Minimal XML serializer for generating OOXML content.
 *
 * Maintains a stack of open elements and serializes to a string. Attribute
 * values are automatically escaped. The output is always valid, well-formed XML.
 */

import { escapeAttribute, escapeHtml } from './html.js';

export class XmlBuilder {
  #parts = [];
  #stack = [];

  /** Open an element, optionally with attributes. */
  el(name, attributes) {
    const attrStr = attributes ? buildAttrs(attributes) : '';
    this.#parts.push(`<${name}${attrStr}>`);
    this.#stack.push(name);
    return this;
  }

  /** Write a self-closing element. */
  self(name, attributes) {
    const attrStr = attributes ? buildAttrs(attributes) : '';
    this.#parts.push(`<${name}${attrStr}/>`);
    return this;
  }

  /** Write text content (automatically escaped). */
  text(value) {
    this.#parts.push(escapeHtml(value));
    return this;
  }

  /** Write raw XML string (caller is responsible for escaping). */
  raw(xml) {
    this.#parts.push(xml);
    return this;
  }

  /** Close the current element. */
  up() {
    const name = this.#stack.pop();
    if (!name) throw new Error('XmlBuilder: up() called with empty stack');
    this.#parts.push(`</${name}>`);
    return this;
  }

  /** Close all open elements and return the final XML string. */
  end() {
    while (this.#stack.length > 0) this.up();
    return this.#parts.join('');
  }

  /** Return current XML without closing open elements (for streaming). */
  toString() {
    return this.#parts.join('');
  }
}

function buildAttrs(attrs) {
  let out = '';
  for (const [key, value] of Object.entries(attrs)) {
    if (value === undefined || value === null || value === false) continue;
    if (value === true) {
      out += ` ${key}="${key}"`;
    } else {
      out += ` ${key}="${escapeAttribute(String(value))}"`;
    }
  }
  return out;
}

/** Convenience: build a complete XML document string with XML declaration. */
export function xmlDoc(body) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
}
