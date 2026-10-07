import { describe, expect, it } from 'vitest';
import { resolveFontFamilyStack, resolveFontFamilyName, EmbeddedFontRegistry } from '../src/fonts/resolver.js';
import { parseSfnt } from '../src/fonts/sfnt-parser.js';
import { buildTestFont } from './helpers/sfnt.js';

describe('resolveFontFamilyStack', () => {
  it('a generic CSS keyword with no registry, or one that never resolves, falls back to Arial', () => {
    expect(resolveFontFamilyStack('sans-serif')).toEqual({ kind: 'fallback', name: 'Arial' });
    expect(resolveFontFamilyStack('serif')).toEqual({ kind: 'fallback', name: 'Arial' });
    expect(resolveFontFamilyStack('monospace')).toEqual({ kind: 'fallback', name: 'Arial' });
  });

  it('a web-safe name with no registered glyph data falls back to Arial just like any other unresolved name', () => {
    expect(resolveFontFamilyStack('Verdana')).toEqual({ kind: 'fallback', name: 'Arial' });
    expect(resolveFontFamilyStack('"TIMES NEW ROMAN"')).toEqual({ kind: 'fallback', name: 'Arial' });
  });

  it('falls back to Arial when nothing in the stack resolves and no registry is given', () => {
    expect(resolveFontFamilyStack('Poppins')).toEqual({ kind: 'fallback', name: 'Arial' });
  });

  it('a real embedded match later in the stack still wins over an earlier name with no registered data', () => {
    const registry = new EmbeddedFontRegistry();
    const font = buildTestFont();
    registry.register('Company Sans', 400, false, parseSfnt(font), font);

    // "Poppins" and "Arial" both have nothing registered — the search keeps
    // going past them rather than stopping early, and finds "Company Sans".
    const result = resolveFontFamilyStack('Poppins, Arial, "Company Sans"', {}, registry);
    expect(result.kind).toBe('embedded');
    expect(result.entry.family).toBe('Company Sans');
  });

  it('resolves to an embedded registry entry when the family/weight/italic combo is registered', () => {
    const registry = new EmbeddedFontRegistry();
    const font = buildTestFont();
    registry.register('Poppins', 700, false, parseSfnt(font), font);

    const result = resolveFontFamilyStack('Poppins', { weight: 700, italic: false }, registry);
    expect(result.kind).toBe('embedded');
    expect(result.entry.family).toBe('Poppins');
    expect(result.entry.weight).toBe(700);
  });

  it('an embedded match earlier in the stack wins over a later built-in name', () => {
    const registry = new EmbeddedFontRegistry();
    const font = buildTestFont();
    registry.register('Poppins', 400, false, parseSfnt(font), font);

    const result = resolveFontFamilyStack('Poppins, Arial', { weight: 400 }, registry);
    expect(result.kind).toBe('embedded');
  });
});

describe('resolveFontFamilyName (DOCX/PPTX)', () => {
  it('maps a bare generic keyword to a real installable font name', () => {
    expect(resolveFontFamilyName('sans-serif')).toBe('Arial');
    expect(resolveFontFamilyName('serif')).toBe('Times New Roman');
    expect(resolveFontFamilyName('monospace')).toBe('Courier New');
  });

  it('passes an arbitrary custom or web-safe name through unchanged, never forcing Arial', () => {
    expect(resolveFontFamilyName('Poppins')).toBe('Poppins');
    expect(resolveFontFamilyName('Georgia')).toBe('Georgia');
  });

  it('passes through undefined/empty unchanged', () => {
    expect(resolveFontFamilyName(undefined)).toBeUndefined();
  });
});

describe('EmbeddedFontRegistry', () => {
  it('match() finds the nearest weight when the exact weight is not registered', () => {
    const registry = new EmbeddedFontRegistry();
    const font = buildTestFont();
    const sfnt = parseSfnt(font);
    registry.register('Poppins', 400, false, sfnt, font);
    registry.register('Poppins', 700, false, sfnt, font);

    // 600 is closer to 700 than to 400.
    expect(registry.match('Poppins', 600, false).weight).toBe(700);
    expect(registry.match('poppins', 500, false).weight).toBe(400);
  });

  it('hasExact() only reports true for an exact weight/italic combo, unlike match()', () => {
    const registry = new EmbeddedFontRegistry();
    const font = buildTestFont();
    registry.register('Poppins', 400, false, parseSfnt(font), font);

    expect(registry.hasExact('Poppins', 400, false)).toBe(true);
    expect(registry.hasExact('Poppins', 700, false)).toBe(false);
    // match() would still find *something* for weight 700 (nearest = 400) —
    // this is exactly the fuzzy-vs-exact distinction hasExact exists for.
    expect(registry.match('Poppins', 700, false)).toBeDefined();
  });

  it('has() is false for a family with nothing registered', () => {
    const registry = new EmbeddedFontRegistry();
    expect(registry.has('Poppins')).toBe(false);
  });
});
