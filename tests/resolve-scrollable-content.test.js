import { describe, expect, it } from 'vitest';
import { parseHtml } from '../src/html-parser/index.js';
import { resolveScrollableContent } from '../src/idm/resolve-scrollable-content.js';
import { convert } from '../src/index.js';

function warnSpy() {
  const calls = [];
  const warn = (code, message) => calls.push({ code, message });
  return { warn, calls };
}

describe('resolveScrollableContent', () => {
  it('preserveScrollableContent:true (default) unwraps the marker exactly like a plain div always has', () => {
    const withMarker = parseHtml('<div style="height:50px; overflow:auto"><p>One</p><p>Two</p></div>');
    const withoutMarker = parseHtml('<div><p>One</p><p>Two</p></div>');
    const { warn, calls } = warnSpy();
    const resolved = resolveScrollableContent(withMarker, {}, warn);
    expect(resolved.children).toEqual(withoutMarker.children);
    expect(calls).toHaveLength(0);
  });

  it('preserveScrollableContent:true is also the default when the option key is entirely absent', () => {
    const doc = parseHtml('<div style="height:50px; overflow:auto"><p>One</p></div>');
    const resolved = resolveScrollableContent(doc, {}, () => {});
    expect(resolved.children[0]).toMatchObject({ type: 'paragraph' });
  });

  it('preserveScrollableContent:false drops children once the declared height budget is exceeded, keeping document order', () => {
    // A large number of short paragraphs, each its own child, against a tiny
    // declared height — comfortably exceeds budget after a few of them.
    const html = '<div style="height:1px; overflow:auto">' + Array.from({ length: 20 }, (_, i) => `<p>Paragraph ${i}</p>`).join('') + '</div>';
    const doc = parseHtml(html);
    const { warn, calls } = warnSpy();
    const resolved = resolveScrollableContent(doc, { preserveScrollableContent: false }, warn);
    expect(resolved.children.length).toBeGreaterThan(0);
    expect(resolved.children.length).toBeLessThan(20);
    // Kept children are a PREFIX of the original, in order.
    resolved.children.forEach((node, i) => {
      expect(node.children[0].value).toBe(`Paragraph ${i}`);
    });
    expect(calls).toHaveLength(1);
    expect(calls[0].code).toBe('content-clipped');
  });

  it('preserveScrollableContent:false always keeps at least the first child, even one alone exceeding the budget', () => {
    const doc = parseHtml('<div style="height:1px; overflow:auto"><p>Only one very long paragraph that alone exceeds the tiny declared height budget by a wide margin.</p></div>');
    const resolved = resolveScrollableContent(doc, { preserveScrollableContent: false }, () => {});
    expect(resolved.children).toHaveLength(1);
    expect(resolved.children[0].type).toBe('paragraph');
  });

  it('preserveScrollableContent:false with nothing dropped emits no warning', () => {
    const doc = parseHtml('<div style="height:5000px; overflow:auto"><p>One</p><p>Two</p></div>');
    const { warn, calls } = warnSpy();
    const resolved = resolveScrollableContent(doc, { preserveScrollableContent: false }, warn);
    expect(resolved.children).toHaveLength(2);
    expect(calls).toHaveLength(0);
  });

  it('resolves a scrollable container nested inside a list item', () => {
    const html = '<ul><li><div style="height:1px; overflow:auto">' + Array.from({ length: 10 }, (_, i) => `<p>Item ${i}</p>`).join('') + '</div></li></ul>';
    const doc = parseHtml(html);
    const resolved = resolveScrollableContent(doc, { preserveScrollableContent: false }, () => {});
    const itemChildren = resolved.children[0].items[0].children;
    expect(itemChildren.length).toBeGreaterThan(0);
    expect(itemChildren.length).toBeLessThan(10);
  });

  it('resolves a scrollable container nested inside a table cell', () => {
    const html = '<table><tr><td><div style="height:1px; overflow:auto">' + Array.from({ length: 10 }, (_, i) => `<p>Cell ${i}</p>`).join('') + '</div></td></tr></table>';
    const doc = parseHtml(html);
    const resolved = resolveScrollableContent(doc, { preserveScrollableContent: false }, () => {});
    const cellChildren = resolved.children[0].rows[0].cells[0].children;
    expect(cellChildren.length).toBeGreaterThan(0);
    expect(cellChildren.length).toBeLessThan(10);
  });
});

describe('convert() end-to-end with preserveScrollableContent', () => {
  const manyParagraphs = Array.from({ length: 30 }, (_, i) => `<p>Paragraph number ${i} with some extra words to take up more line-wrapped height.</p>`).join('');
  const html = `<div style="height:20px; overflow:auto">${manyParagraphs}</div>`;

  it('defaults to preserving all content when the option is omitted', async () => {
    const result = await convert(html, { format: 'markdown' });
    for (let i = 0; i < 30; i++) {
      expect(result.buffer.toString('utf8')).toContain(`Paragraph number ${i}`);
    }
    expect(result.warnings.some((w) => w.code === 'content-clipped')).toBe(false);
  });

  it('drops content and warns when preserveScrollableContent:false', async () => {
    const result = await convert(html, { format: 'markdown', preserveScrollableContent: false });
    const text = result.buffer.toString('utf8');
    expect(text).toContain('Paragraph number 0');
    expect(text).not.toContain('Paragraph number 29');
    expect(result.warnings.some((w) => w.code === 'content-clipped')).toBe(true);
  });

  it('applies consistently across every binary/text format', async () => {
    for (const format of ['markdown', 'docx', 'xlsx', 'pptx', 'pdf']) {
      const preserved = await convert(html, { format });
      const clipped = await convert(html, { format, preserveScrollableContent: false });
      expect(clipped.buffer.length).toBeGreaterThan(0);
      expect(preserved.buffer.length).toBeGreaterThan(0);
    }
  });
});
