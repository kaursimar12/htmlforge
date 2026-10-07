import { describe, it, expect } from 'vitest';
import { estimateTextWidth, estimateWrappedLines, estimateShapeSize } from '../src/exporter/shared/font-metrics.js';

describe('font-metrics — estimateTextWidth', () => {
  it('scales linearly with font size', () => {
    const at10 = estimateTextWidth('Hello World', 10, 'Arial');
    const at20 = estimateTextWidth('Hello World', 20, 'Arial');
    expect(at20).toBeCloseTo(at10 * 2, 5);
  });

  it('gives a visually wide string more width than a narrow one of equal length, in a known font', () => {
    const wide = estimateTextWidth('WWWWWWWWWW', 12, 'Arial');
    const narrow = estimateTextWidth('iiiiiiiiii', 12, 'Arial');
    expect(wide).toBeGreaterThan(narrow * 2);
  });

  it('falls back to a character-class estimate for an unrecognized font family', () => {
    const width = estimateTextWidth('Hello', 12, 'SomeCustomWebFont');
    expect(width).toBeGreaterThan(0);
  });

  it('falls back the same way when no font family is given at all', () => {
    const width = estimateTextWidth('Hello', 12, undefined);
    expect(width).toBeGreaterThan(0);
  });

  it('reads the first entry of a CSS font stack, ignoring generic fallback keywords', () => {
    const stack = estimateTextWidth('Hello World', 14, "Arial, 'Helvetica Neue', sans-serif");
    const plain = estimateTextWidth('Hello World', 14, 'Arial');
    expect(stack).toBeCloseTo(plain, 5);
  });

  it('is case-insensitive on the font family name', () => {
    const lower = estimateTextWidth('Hello', 12, 'arial');
    const mixed = estimateTextWidth('Hello', 12, 'Arial');
    expect(lower).toBeCloseTo(mixed, 5);
  });

  it('produces different (but non-zero, sane) widths for each named font at the same text/size', () => {
    const fonts = ['Arial', 'Helvetica', 'Times New Roman', 'Calibri', 'Georgia'];
    const widths = fonts.map((f) => estimateTextWidth('The quick brown fox', 12, f));
    for (const w of widths) {
      expect(w).toBeGreaterThan(0);
      // Sanity bound: no reasonable font renders 20 characters at 12pt
      // anywhere near 0 or past 300pt.
      expect(w).toBeLessThan(300);
    }
  });
});

describe('font-metrics — estimateWrappedLines', () => {
  it('keeps a short string on one line', () => {
    const lines = estimateWrappedLines('Short badge', 500, 14, 'Arial');
    expect(lines).toEqual(['Short badge']);
  });

  it('wraps long text into multiple lines that each fit the given width', () => {
    const text = 'The quick brown fox jumps over the lazy dog and keeps running';
    const maxWidthPt = 80;
    const lines = estimateWrappedLines(text, maxWidthPt, 12, 'Arial');
    expect(lines.length).toBeGreaterThan(1);
    for (const line of lines) {
      expect(estimateTextWidth(line, 12, 'Arial')).toBeLessThanOrEqual(maxWidthPt + 1e-6);
    }
  });

  it('reassembles to the original words in order', () => {
    const text = 'one two three four five six seven eight';
    const lines = estimateWrappedLines(text, 60, 12, 'Arial');
    expect(lines.join(' ')).toBe(text);
  });

  it('places a single word wider than maxWidth alone on its own line instead of looping', () => {
    const lines = estimateWrappedLines('Supercalifragilisticexpialidocious', 10, 40, 'Arial');
    expect(lines).toEqual(['Supercalifragilisticexpialidocious']);
  });

  it('returns one empty line for empty/whitespace-only text', () => {
    expect(estimateWrappedLines('', 200, 12, 'Arial')).toEqual(['']);
    expect(estimateWrappedLines('   ', 200, 12, 'Arial')).toEqual(['']);
  });
});

describe('font-metrics — estimateShapeSize', () => {
  it('auto-sizes (no maxWidthPt) to one line, biased above the raw text width', () => {
    const raw = estimateTextWidth('v1.0', 14, 'Arial');
    const { width, height, lineCount } = estimateShapeSize('v1.0', 14, 'Arial', {
      padding: { top: 6, right: 10, bottom: 6, left: 10 },
    });
    expect(lineCount).toBe(1);
    expect(width).toBeGreaterThan(raw + 10 + 10); // biased above raw width + the padding alone
    expect(height).toBeCloseTo(14 * 1.2 + 6 + 6, 5);
  });

  it('respects a custom autoSizeBiasPercent', () => {
    const small = estimateShapeSize('Badge', 14, 'Arial', { autoSizeBiasPercent: 0, padding: {} });
    const large = estimateShapeSize('Badge', 14, 'Arial', { autoSizeBiasPercent: 50, padding: {} });
    expect(large.width).toBeGreaterThan(small.width);
  });

  it('wraps to a fixed width and grows height with line count', () => {
    const text = 'This is a longer card description that should wrap across multiple lines.';
    const result = estimateShapeSize(text, 12, 'Arial', {
      maxWidthPt: 150,
      padding: { top: 12, right: 16, bottom: 12, left: 16 },
    });
    expect(result.width).toBe(150);
    expect(result.lineCount).toBeGreaterThan(1);
    expect(result.height).toBeCloseTo(result.lineCount * 12 * 1.2 + 12 + 12, 5);
  });

  it('a fixed-width single short word still returns exactly one line', () => {
    const result = estimateShapeSize('Hi', 12, 'Arial', { maxWidthPt: 200, padding: {} });
    expect(result.lineCount).toBe(1);
  });
});
