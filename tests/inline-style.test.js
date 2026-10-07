import { describe, it, expect } from 'vitest';
import {
  parseInlineStyleAttr,
  extractRunStyle,
  extractAlign,
  extractHeightPx,
  extractMaxHeightPx,
  extractOverflow,
  extractOverflowX,
  extractOverflowY,
  extractBackgroundColor,
  extractTableBorder,
  extractCellBorder,
} from '../src/html-parser/inline-style.js';

describe('parseInlineStyleAttr', () => {
  it('splits declarations on ; and properties on the first :', () => {
    expect(parseInlineStyleAttr('color: red; font-weight:bold')).toEqual({
      color: 'red',
      'font-weight': 'bold',
    });
  });

  it('lowercases the property name but preserves the value case', () => {
    expect(parseInlineStyleAttr('COLOR: Red')).toEqual({ color: 'Red' });
  });

  it('ignores a declaration with no colon', () => {
    expect(parseInlineStyleAttr('color: red; garbage; font-style:italic')).toEqual({
      color: 'red',
      'font-style': 'italic',
    });
  });

  it('ignores an empty property or empty value', () => {
    expect(parseInlineStyleAttr(': red; color: ;font-weight:bold')).toEqual({ 'font-weight': 'bold' });
  });

  it('returns {} for undefined/empty input', () => {
    expect(parseInlineStyleAttr(undefined)).toEqual({});
    expect(parseInlineStyleAttr('')).toEqual({});
  });

  it('a value containing a colon (e.g. a url()) keeps everything after the first colon', () => {
    expect(parseInlineStyleAttr('background:url(http://x.com/a.png)')).toEqual({
      background: 'url(http://x.com/a.png)',
    });
  });
});

describe('extractRunStyle', () => {
  it('returns undefined when no recognized property is present', () => {
    expect(extractRunStyle('margin:10px; padding:5px')).toBeUndefined();
  });

  it('returns undefined for an empty/missing style attribute', () => {
    expect(extractRunStyle(undefined)).toBeUndefined();
    expect(extractRunStyle('')).toBeUndefined();
  });

  it('parses a 6-digit hex color, uppercased', () => {
    expect(extractRunStyle('color:#ff0000')).toEqual({ color: 'FF0000' });
  });

  it('parses a 3-digit hex color, expanded and uppercased', () => {
    expect(extractRunStyle('color:#f00')).toEqual({ color: 'FF0000' });
  });

  it('parses an rgb() color', () => {
    expect(extractRunStyle('color:rgb(255, 0, 0)')).toEqual({ color: 'FF0000' });
  });

  it('parses an rgba() color, ignoring the alpha channel', () => {
    expect(extractRunStyle('color:rgba(255, 0, 0, 0.5)')).toEqual({ color: 'FF0000' });
  });

  it('parses a named color', () => {
    expect(extractRunStyle('color:blue')).toEqual({ color: '0000FF' });
  });

  it('is case-insensitive for named colors and hex', () => {
    expect(extractRunStyle('color:BLUE')).toEqual({ color: '0000FF' });
    expect(extractRunStyle('color:#FF0000')).toEqual({ color: 'FF0000' });
  });

  it('drops an unrecognized color value', () => {
    expect(extractRunStyle('color:currentColor')).toBeUndefined();
  });

  it('parses background-color the same way as color', () => {
    expect(extractRunStyle('background-color:#00ff00')).toEqual({ backgroundColor: '00FF00' });
  });

  it('falls back to the background shorthand when background-color is absent', () => {
    expect(extractRunStyle('background:yellow')).toEqual({ backgroundColor: 'FFFF00' });
  });

  it('prefers background-color over the background shorthand when both are present', () => {
    expect(extractRunStyle('background:yellow; background-color:#0000ff')).toEqual({ backgroundColor: '0000FF' });
  });

  it('a background shorthand with a non-color value (e.g. an image) does not resolve', () => {
    expect(extractRunStyle('background:url(x.png) no-repeat')).toBeUndefined();
  });

  it.each([
    ['bold', 700], ['bolder', 700], ['600', 600], ['700', 700], ['900', 900],
  ])('treats font-weight:%s as bold', (weight, fontWeight) => {
    expect(extractRunStyle(`font-weight:${weight}`)).toEqual({ bold: true, fontWeight });
  });

  it.each([
    ['normal', 400], ['400', 400], ['500', 500], ['399', 399],
  ])('does not treat font-weight:%s as bold, but still carries the numeric weight', (weight, fontWeight) => {
    expect(extractRunStyle(`font-weight:${weight}`)).toEqual({ fontWeight });
  });

  it.each(['italic', 'oblique'])('treats font-style:%s as italic', (value) => {
    expect(extractRunStyle(`font-style:${value}`)).toEqual({ italic: true });
  });

  it('does not treat font-style:normal as italic', () => {
    expect(extractRunStyle('font-style:normal')).toBeUndefined();
  });

  it('parses text-decoration:underline', () => {
    expect(extractRunStyle('text-decoration:underline')).toEqual({ underline: true });
  });

  it('parses text-decoration:line-through', () => {
    expect(extractRunStyle('text-decoration:line-through')).toEqual({ strikethrough: true });
  });

  it('parses both underline and line-through from one shorthand value', () => {
    expect(extractRunStyle('text-decoration:underline line-through')).toEqual({
      underline: true,
      strikethrough: true,
    });
  });

  it('falls back to text-decoration-line when text-decoration is absent', () => {
    expect(extractRunStyle('text-decoration-line:underline')).toEqual({ underline: true });
  });

  it('parses font-size in px, converted to pt', () => {
    expect(extractRunStyle('font-size:16px')).toEqual({ fontSizePt: 12 });
  });

  it('parses font-size in pt, unconverted', () => {
    expect(extractRunStyle('font-size:14pt')).toEqual({ fontSizePt: 14 });
  });

  it('ignores font-size in a relative unit (em/%) — no inherited base size to resolve it against', () => {
    expect(extractRunStyle('font-size:1.5em')).toBeUndefined();
    expect(extractRunStyle('font-size:150%')).toBeUndefined();
  });

  it.each([
    ['xx-small', 6.75], ['x-small', 7.5], ['small', 9.75], ['medium', 12],
    ['large', 13.5], ['x-large', 18], ['xx-large', 24], ['xxx-large', 36],
  ])('resolves the CSS absolute-size keyword font-size:%s to its fixed pt value', (keyword, pt) => {
    expect(extractRunStyle(`font-size:${keyword}`)).toEqual({ fontSizePt: pt });
  });

  it('is case-insensitive for the absolute-size keywords', () => {
    expect(extractRunStyle('font-size:LARGE')).toEqual({ fontSizePt: 13.5 });
  });

  it('combines every recognized property from one style attribute', () => {
    const style = extractRunStyle(
      'color:#ff0000;background-color:#ffff00;font-weight:bold;font-style:italic;text-decoration:underline;font-size:16px',
    );
    expect(style).toEqual({
      color: 'FF0000',
      backgroundColor: 'FFFF00',
      bold: true,
      fontWeight: 700,
      italic: true,
      underline: true,
      fontSizePt: 12,
    });
  });

  it('parses a simple font-family value', () => {
    expect(extractRunStyle('font-family:Arial')).toEqual({ fontFamily: 'Arial' });
  });

  it('uses only the first font in a comma-separated fallback stack', () => {
    expect(extractRunStyle("font-family:'Times New Roman', serif")).toEqual({ fontFamily: 'Times New Roman' });
  });

  it('strips surrounding quotes (single or double) from the font name', () => {
    expect(extractRunStyle('font-family:"Courier New"')).toEqual({ fontFamily: 'Courier New' });
    expect(extractRunStyle("font-family:'Courier New'")).toEqual({ fontFamily: 'Courier New' });
  });

  it('a bare unquoted multi-word font name is used as-is (no comma to split on)', () => {
    expect(extractRunStyle('font-family:Times New Roman')).toEqual({ fontFamily: 'Times New Roman' });
  });
});

describe('extractAlign', () => {
  it.each(['left', 'center', 'right', 'justify'])('accepts text-align:%s', (align) => {
    expect(extractAlign(`text-align:${align}`)).toBe(align);
  });

  it('is case-insensitive', () => {
    expect(extractAlign('text-align:CENTER')).toBe('center');
  });

  it('rejects an unrecognized alignment value', () => {
    expect(extractAlign('text-align:start')).toBeUndefined();
  });

  it('returns undefined when text-align is absent', () => {
    expect(extractAlign('color:red')).toBeUndefined();
    expect(extractAlign(undefined)).toBeUndefined();
  });
});

describe('extractHeightPx', () => {
  it('parses an integer px height', () => {
    expect(extractHeightPx('height:4px')).toBe(4);
  });

  it('parses a decimal px height', () => {
    expect(extractHeightPx('height:4.5px')).toBe(4.5);
  });

  it('ignores a non-px unit (pt/%/em)', () => {
    expect(extractHeightPx('height:10pt')).toBeUndefined();
    expect(extractHeightPx('height:50%')).toBeUndefined();
    expect(extractHeightPx('height:2em')).toBeUndefined();
  });

  it('ignores a bare unitless number (not valid CSS anyway)', () => {
    expect(extractHeightPx('height:4')).toBeUndefined();
  });

  it('returns undefined when height is absent', () => {
    expect(extractHeightPx('width:100px')).toBeUndefined();
    expect(extractHeightPx(undefined)).toBeUndefined();
  });
});

describe('extractMaxHeightPx', () => {
  it('parses an integer px max-height', () => {
    expect(extractMaxHeightPx('max-height:50px')).toBe(50);
  });

  it('ignores a non-px unit (pt/%/em)', () => {
    expect(extractMaxHeightPx('max-height:50%')).toBeUndefined();
    expect(extractMaxHeightPx('max-height:2em')).toBeUndefined();
  });

  it('returns undefined when max-height is absent', () => {
    expect(extractMaxHeightPx('height:50px')).toBeUndefined();
    expect(extractMaxHeightPx(undefined)).toBeUndefined();
  });
});

describe('extractOverflow/extractOverflowX/extractOverflowY', () => {
  it('recognizes visible/hidden/scroll/auto', () => {
    for (const value of ['visible', 'hidden', 'scroll', 'auto']) {
      expect(extractOverflow(`overflow:${value}`)).toBe(value);
    }
  });

  it('is case-insensitive', () => {
    expect(extractOverflow('overflow:Auto')).toBe('auto');
  });

  it('returns undefined for an unrecognized value or when absent', () => {
    expect(extractOverflow('overflow:clip')).toBeUndefined();
    expect(extractOverflow(undefined)).toBeUndefined();
  });

  it('overflow-x/overflow-y read their own longhand first', () => {
    expect(extractOverflowX('overflow-x:scroll; overflow-y:hidden')).toBe('scroll');
    expect(extractOverflowY('overflow-x:scroll; overflow-y:hidden')).toBe('hidden');
  });

  it('overflow-x/overflow-y fall back to the shorthand when their own longhand is absent', () => {
    expect(extractOverflowX('overflow:auto')).toBe('auto');
    expect(extractOverflowY('overflow:auto')).toBe('auto');
  });

  it('longhand wins over shorthand when both are declared', () => {
    expect(extractOverflowX('overflow:visible; overflow-x:hidden')).toBe('hidden');
  });
});

describe('extractBackgroundColor', () => {
  it('parses a hex background-color', () => {
    expect(extractBackgroundColor('background-color:#2563eb')).toBe('2563EB');
  });

  it('parses a named background-color', () => {
    expect(extractBackgroundColor('background-color:orange')).toBe('FFA500');
  });

  it('returns undefined when background-color is absent or unrecognized', () => {
    expect(extractBackgroundColor('color:red')).toBeUndefined();
    expect(extractBackgroundColor('background-color:transparent')).toBeUndefined();
  });

  it('falls back to the background shorthand when background-color is absent', () => {
    expect(extractBackgroundColor('background:#2563eb')).toBe('2563EB');
  });
});

describe('extractTableBorder', () => {
  it('parses the border shorthand regardless of token order', () => {
    expect(extractTableBorder('border:1px solid #e5e7eb')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
    expect(extractTableBorder('border:solid #e5e7eb 1px')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
    expect(extractTableBorder('border:#e5e7eb 1px solid')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
  });

  it('longhand border-width/style/color override the shorthand', () => {
    const style = extractTableBorder('border:1px solid red; border-color:#0000ff; border-width:2px');
    expect(style).toEqual({ color: '0000FF', sizeEighthPt: 12, style: 'single' });
  });

  it('border-style:none/hidden suppresses the border entirely', () => {
    expect(extractTableBorder('border:1px solid #e5e7eb; border-style:none')).toBeUndefined();
    expect(extractTableBorder('border-style:hidden')).toBeUndefined();
  });

  it('a color with no width still resolves, using the default hairline width', () => {
    expect(extractTableBorder('border-color:#e5e7eb')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
  });

  it('a width with no color still resolves, defaulting to black', () => {
    expect(extractTableBorder('border-width:2px')).toEqual({ color: '000000', sizeEighthPt: 12, style: 'single' });
  });

  it('returns undefined when no border is declared at all', () => {
    expect(extractTableBorder('padding:10px')).toBeUndefined();
    expect(extractTableBorder(undefined)).toBeUndefined();
  });

  it('resolves the full CSS extended named-color set, not just the dozen most common ones', () => {
    expect(extractTableBorder('border:5px double crimson')).toEqual({ color: 'DC143C', sizeEighthPt: 30, style: 'double' });
    expect(extractTableBorder('border-color:rebeccapurple')).toEqual({ color: '663399', sizeEighthPt: 6, style: 'single' });
    expect(extractTableBorder('border-color:cornflowerblue')).toEqual({ color: '6495ED', sizeEighthPt: 6, style: 'single' });
  });

  it('a very thin width is floored at the minimum 2-eighths-of-a-point border', () => {
    expect(extractTableBorder('border:0.1px solid #000')).toEqual({ color: '000000', sizeEighthPt: 2, style: 'single' });
  });

  // border-width:medium is a real CSS keyword (~3px) — it must not be
  // misinterpreted via the font-size keyword table, where "medium" means
  // 16px (12pt) instead. An unrecognized width keyword falls back to the
  // same hairline default as any other unresolvable width.
  it('a border-width keyword (medium/thin/thick) is not confused with the font-size keyword of the same name', () => {
    expect(extractTableBorder('border-width:medium; border-color:#000')).toEqual({ color: '000000', sizeEighthPt: 6, style: 'single' });
  });
});

describe('extractCellBorder', () => {
  it('parses a single side (e.g. border-bottom) independent of the others', () => {
    const style = 'border-bottom:1px solid #e5e7eb';
    expect(extractCellBorder(style, 'bottom')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
    expect(extractCellBorder(style, 'top')).toBeUndefined();
    expect(extractCellBorder(style, 'left')).toBeUndefined();
    expect(extractCellBorder(style, 'right')).toBeUndefined();
  });

  it('each side is independently declarable on the same element', () => {
    const style = 'border-top:1px solid red; border-bottom:2px solid blue';
    expect(extractCellBorder(style, 'top')).toEqual({ color: 'FF0000', sizeEighthPt: 6, style: 'single' });
    expect(extractCellBorder(style, 'bottom')).toEqual({ color: '0000FF', sizeEighthPt: 12, style: 'single' });
  });

  it('longhand border-{side}-color/width overrides the border-{side} shorthand', () => {
    const style = 'border-bottom:1px solid red; border-bottom-color:#00ff00';
    expect(extractCellBorder(style, 'bottom')).toEqual({ color: '00FF00', sizeEighthPt: 6, style: 'single' });
  });

  it('border-{side}-style:none suppresses just that side', () => {
    expect(extractCellBorder('border-bottom:1px solid red; border-bottom-style:none', 'bottom')).toBeUndefined();
  });

  it('falls back to the plain `border` shorthand applied to all four sides', () => {
    const style = 'border:1px solid #e5e7eb';
    expect(extractCellBorder(style, 'bottom')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
    expect(extractCellBorder(style, 'top')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
    expect(extractCellBorder(style, 'left')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
    expect(extractCellBorder(style, 'right')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
  });

  it('a per-side declaration overrides the all-sides `border` shorthand for that side only', () => {
    const style = 'border:1px solid #e5e7eb; border-bottom:2px solid red';
    expect(extractCellBorder(style, 'bottom')).toEqual({ color: 'FF0000', sizeEighthPt: 12, style: 'single' });
    expect(extractCellBorder(style, 'top')).toEqual({ color: 'E5E7EB', sizeEighthPt: 6, style: 'single' });
  });

  it('returns undefined when that side has no border declared', () => {
    expect(extractCellBorder('color:red', 'bottom')).toBeUndefined();
    expect(extractCellBorder(undefined, 'bottom')).toBeUndefined();
  });
});
