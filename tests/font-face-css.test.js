import { describe, expect, it } from 'vitest';
import { parseFontFaceRules, parseImports, unicodeRangeCovers } from '../src/fonts/font-face-css.js';

describe('parseFontFaceRules', () => {
  it('parses family/weight/style and one src url()', () => {
    const css = `
      @font-face {
        font-family: 'Poppins';
        font-weight: 700;
        font-style: normal;
        src: url(https://fonts.gstatic.com/s/poppins/v23/bold.woff2) format('woff2');
      }
    `;
    const rules = parseFontFaceRules(css);
    expect(rules).toEqual([
      { family: 'Poppins', weight: 700, italic: false, srcs: [{ url: 'https://fonts.gstatic.com/s/poppins/v23/bold.woff2', format: 'woff2' }] },
    ]);
  });

  it('parses multiple src entries in declared order (woff2 then ttf fallback)', () => {
    const css = `@font-face {
      font-family: Inter;
      src: url(a.woff2) format('woff2'), url(a.ttf) format('truetype');
    }`;
    const [rule] = parseFontFaceRules(css);
    expect(rule.srcs).toEqual([
      { url: 'a.woff2', format: 'woff2' },
      { url: 'a.ttf', format: 'truetype' },
    ]);
  });

  it('defaults weight to 400 and italic to false when unspecified', () => {
    const [rule] = parseFontFaceRules('@font-face { font-family: Roboto; src: url(r.woff2); }');
    expect(rule.weight).toBe(400);
    expect(rule.italic).toBe(false);
  });

  it('recognizes italic/oblique font-style', () => {
    const css = '@font-face { font-family: Lato; font-style: italic; src: url(l.woff2); }';
    expect(parseFontFaceRules(css)[0].italic).toBe(true);
  });

  it('parses several @font-face rules (different weights of the same family) from one stylesheet', () => {
    const css = `
      @font-face { font-family: Poppins; font-weight: 400; src: url(regular.woff2); }
      @font-face { font-family: Poppins; font-weight: 700; src: url(bold.woff2); }
    `;
    const rules = parseFontFaceRules(css);
    expect(rules).toHaveLength(2);
    expect(rules.map((r) => r.weight)).toEqual([400, 700]);
  });

  it('skips a rule with no usable src', () => {
    const css = '@font-face { font-family: Broken; }';
    expect(parseFontFaceRules(css)).toEqual([]);
  });

  it('ignores plain selector rules and only matches @font-face blocks', () => {
    const css = '.title { color: red; } @font-face { font-family: X; src: url(x.woff2); }';
    expect(parseFontFaceRules(css)).toHaveLength(1);
  });
});

describe('unicode-range parsing', () => {
  it('parses a real Google/Bunny-style "latin" subset rule, and covers Basic Latin', () => {
    const css = `@font-face {
      font-family: Inter;
      font-weight: 400;
      src: url(inter-latin.woff2) format('woff2');
      unicode-range: U+0000-00FF,U+0131,U+2000-206F,U+2212;
    }`;
    const [rule] = parseFontFaceRules(css);
    expect(unicodeRangeCovers(rule.unicodeRange, 0x41)).toBe(true); // 'A'
    expect(unicodeRangeCovers(rule.unicodeRange, 0x61)).toBe(true); // 'a'
    expect(unicodeRangeCovers(rule.unicodeRange, 0x400)).toBe(false); // Cyrillic
  });

  it('a "cyrillic" subset rule does not cover Basic Latin', () => {
    const css = `@font-face {
      font-family: Inter;
      font-weight: 400;
      src: url(inter-cyrillic.woff2);
      unicode-range: U+0400-045F,U+2116;
    }`;
    const [rule] = parseFontFaceRules(css);
    expect(unicodeRangeCovers(rule.unicodeRange, 0x41)).toBe(false);
    expect(unicodeRangeCovers(rule.unicodeRange, 0x420)).toBe(true);
  });

  it('expands a hex-wildcard range (U+4E??)', () => {
    const css = "@font-face { font-family: X; src: url(x.woff2); unicode-range: U+4E??; }";
    const [rule] = parseFontFaceRules(css);
    expect(unicodeRangeCovers(rule.unicodeRange, 0x4e00)).toBe(true);
    expect(unicodeRangeCovers(rule.unicodeRange, 0x4eff)).toBe(true);
    expect(unicodeRangeCovers(rule.unicodeRange, 0x4d00)).toBe(false);
  });

  it('no unicode-range declared at all means "covers everything"', () => {
    const [rule] = parseFontFaceRules('@font-face { font-family: X; src: url(x.woff2); }');
    expect(rule.unicodeRange).toBeUndefined();
    expect(unicodeRangeCovers(rule.unicodeRange, 0x41)).toBe(true);
    expect(unicodeRangeCovers(rule.unicodeRange, 0x1f600)).toBe(true);
  });
});

describe('parseImports', () => {
  it('parses @import url(...) with and without quotes', () => {
    expect(parseImports('@import url(fonts.css);')).toEqual(['fonts.css']);
    expect(parseImports('@import url("fonts.css");')).toEqual(['fonts.css']);
    expect(parseImports("@import url('fonts.css');")).toEqual(['fonts.css']);
  });

  it('parses the bare-string @import form', () => {
    expect(parseImports('@import "fonts.css";')).toEqual(['fonts.css']);
  });

  it('parses multiple imports from one stylesheet', () => {
    const css = '@import url(a.css); @import url(b.css);';
    expect(parseImports(css)).toEqual(['a.css', 'b.css']);
  });

  it('returns an empty array when there are no imports', () => {
    expect(parseImports('.title { color: red; }')).toEqual([]);
  });
});
