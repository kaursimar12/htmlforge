import { describe, expect, it } from 'vitest';
import {
  attr,
  childElements,
  decodeEntities,
  find,
  findAll,
  firstChild,
  parseXml,
  textContent,
} from '../src/utils/xml.js';

describe('decodeEntities', () => {
  it('decodes named and numeric references', () => {
    expect(decodeEntities('a &amp; b &lt;c&gt; &#65; &#x42;')).toBe('a & b <c> A B');
  });
  it('leaves unknown entities untouched', () => {
    expect(decodeEntities('&nbsp;x')).toBe('&nbsp;x');
  });
});

describe('parseXml', () => {
  it('parses elements, attributes, and text', () => {
    const root = parseXml('<w:p w:rsid="1"><w:r><w:t>Hello</w:t></w:r></w:p>');
    expect(root.name).toBe('w:p');
    expect(attr(root, 'w:rsid')).toBe('1');
    expect(textContent(root)).toBe('Hello');
  });

  it('matches helpers on prefixed or local names', () => {
    const root = parseXml('<w:body><w:p><w:t>A</w:t></w:p><w:p><w:t>B</w:t></w:p></w:body>');
    expect(childElements(root, 'p')).toHaveLength(2);
    expect(childElements(root, 'w:p')).toHaveLength(2);
    expect(findAll(root, 't').map(textContent)).toEqual(['A', 'B']);
    expect(textContent(find(root, 'p'))).toBe('A');
  });

  it('handles self-closing tags, CDATA, comments, and declarations', () => {
    const root = parseXml(
      '<?xml version="1.0"?><!-- c --><root><br/><![CDATA[<raw> & ]]><x>y</x></root>',
    );
    expect(root.name).toBe('root');
    expect(firstChild(root, 'br')).toBeDefined();
    expect(textContent(root)).toBe('<raw> & y');
  });

  it('respects xml:space-style preserved whitespace in text', () => {
    const root = parseXml('<t>  spaced  </t>');
    expect(textContent(root)).toBe('  spaced  ');
  });

  it('is tolerant of a missing close tag', () => {
    const root = parseXml('<a><b>text</a>');
    expect(textContent(root)).toBe('text');
  });
});
