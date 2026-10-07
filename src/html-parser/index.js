/**
 * HTML Parser — public entry point.
 *
 * HTML string → IDM DocumentNode, ready for any exporter to consume.
 */

import { buildTree } from './tree-builder.js';
import { domToIdm } from './dom-to-idm.js';

export function parseHtml(html) {
  const dom = buildTree(html);
  return domToIdm(dom);
}

export { buildTree } from './tree-builder.js';
export { domToIdm } from './dom-to-idm.js';
export { tokenize } from './tokenizer.js';
