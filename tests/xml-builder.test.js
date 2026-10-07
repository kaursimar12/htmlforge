import { describe, it, expect } from 'vitest';
import { XmlBuilder, xmlDoc } from '../src/utils/xml-builder.js';
import { parseXml, attr, textContent, childElements, find } from '../src/utils/xml.js';

describe('XmlBuilder', () => {
  it('round-trips a simple element through parseXml()', () => {
    const xml = new XmlBuilder().el('root').text('hello').end();
    const el = parseXml(xml);
    expect(el.name).toBe('root');
    expect(textContent(el)).toBe('hello');
  });

  it('round-trips namespace prefixes', () => {
    const xml = new XmlBuilder()
      .el('w:document')
        .el('w:body')
          .el('w:p')
            .el('w:r')
              .el('w:t').text('Hello').up()
            .up()
          .up()
        .up()
      .end();

    const doc = parseXml(xml);
    expect(doc.name).toBe('w:document');
    const t = find(doc, 't');
    expect(t).toBeDefined();
    expect(textContent(t)).toBe('Hello');
  });

  it('round-trips attributes', () => {
    const xml = new XmlBuilder()
      .el('w:p', { 'w:rsid': 'ABC123', 'w:rsidR': '00FF1234' })
      .end();

    const el = parseXml(xml);
    expect(attr(el, 'w:rsid')).toBe('ABC123');
    expect(attr(el, 'w:rsidR')).toBe('00FF1234');
  });

  it('escapes text content', () => {
    const xml = new XmlBuilder().el('t').text('<b>a & b</b>').end();
    const el = parseXml(xml);
    expect(textContent(el)).toBe('<b>a & b</b>');
  });

  it('escapes attribute values', () => {
    const xml = new XmlBuilder().el('a', { href: 'url?a=1&b=2' }).end();
    const el = parseXml(xml);
    expect(attr(el, 'href')).toBe('url?a=1&b=2');
  });

  it('self() writes a self-closing element', () => {
    const xml = new XmlBuilder()
      .el('root')
        .self('w:br')
        .self('w:lastRenderedPageBreak')
      .end();

    const el = parseXml(xml);
    const children = childElements(el);
    expect(children).toHaveLength(2);
    expect(children[0].name).toBe('w:br');
    expect(children[1].name).toBe('w:lastRenderedPageBreak');
  });

  it('omits undefined/null/false attributes', () => {
    const xml = new XmlBuilder()
      .el('w:r', { keep: 'yes', skip: undefined, also: null, falsy: false })
      .end();

    const el = parseXml(xml);
    expect(attr(el, 'keep')).toBe('yes');
    expect(attr(el, 'skip')).toBeUndefined();
    expect(attr(el, 'also')).toBeUndefined();
    expect(attr(el, 'falsy')).toBeUndefined();
  });

  it('end() closes all unclosed elements', () => {
    const xml = new XmlBuilder().el('a').el('b').el('c').end();
    const el = parseXml(xml);
    expect(el.name).toBe('a');
    expect(find(el, 'b')).toBeDefined();
    expect(find(el, 'c')).toBeDefined();
  });

  it('up() throws when stack is empty', () => {
    expect(() => new XmlBuilder().up()).toThrow();
  });
});

describe('xmlDoc', () => {
  it('prepends the XML declaration', () => {
    const xml = xmlDoc('<root/>');
    expect(xml).toMatch(/^<\?xml version="1\.0"/);
    expect(xml).toContain('<root/>');
  });
});
