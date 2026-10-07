/**
 * Regression coverage for a real bug found while live-testing this feature
 * against actual Google Fonts/Bunny Fonts stylesheets (see
 * dev/inline-style-sample.html): providers split ONE family+weight across
 * several @font-face rules by Unicode subset (`latin`, `cyrillic`, `greek`,
 * ...), each with different glyph coverage. Fetching every subset
 * concurrently and letting whichever finishes last win the registration is
 * a race — it can silently embed a font missing ordinary Latin letters.
 * stylesheet-resolver.js now groups by family|weight|italic and picks the
 * rule that actually covers Basic Latin before fetching anything.
 */
import { describe, it, expect, afterEach, vi } from 'vitest';
import { EmbeddedFontRegistry } from '../src/fonts/resolver.js';
import { discoverFontsFromDocument, clearStylesheetCaches } from '../src/fonts/stylesheet-resolver.js';
import { buildTestFont } from './helpers/sfnt.js';

describe('discoverFontsFromDocument — multi-subset @font-face resolution', () => {
  afterEach(() => { vi.unstubAllGlobals(); clearStylesheetCaches(); });

  const LATIN_FONT = buildTestFont({ glyphs: [{ cp: 0x41, width: 500 }, { cp: 0x61, width: 500 }] }); // A, a
  const CYRILLIC_ONLY_FONT = buildTestFont({ glyphs: [{ cp: 0x410, width: 500 }] }); // А (Cyrillic A), no Latin at all

  const STYLESHEET_CSS = `
    @font-face {
      font-family: 'TestFont';
      font-weight: 400;
      src: url(https://example.com/testfont-cyrillic.ttf);
      unicode-range: U+0400-04FF;
    }
    @font-face {
      font-family: 'TestFont';
      font-weight: 400;
      src: url(https://example.com/testfont-latin.ttf);
      unicode-range: U+0000-00FF;
    }
  `;

  function mockFetch() {
    const fn = vi.fn(async (url) => {
      const body = url.includes('cyrillic') ? CYRILLIC_ONLY_FONT
        : url.includes('latin') ? LATIN_FONT
        : Buffer.from(STYLESHEET_CSS);
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        headers: new Map(),
        text: async () => STYLESHEET_CSS,
        arrayBuffer: async () => body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength),
      };
    });
    vi.stubGlobal('fetch', fn);
    return fn;
  }

  function documentUsing(family) {
    return { metadata: { fontFamily: family, stylesheetLinks: ['https://example.com/fonts.css'] }, children: [] };
  }

  it('picks the subset rule that covers Basic Latin, not whichever fetch happens to finish last', async () => {
    mockFetch();
    const registry = new EmbeddedFontRegistry();
    await discoverFontsFromDocument(documentUsing('TestFont'), {}, undefined, registry);

    const entry = registry.match('TestFont', 400, false);
    expect(entry).toBeDefined();
    expect(entry.sfnt.glyphIdForCodepoint(0x41)).toBeDefined(); // 'A'
    expect(entry.sfnt.glyphIdForCodepoint(0x61)).toBeDefined(); // 'a'
  });

  it('only fetches the one font FILE it selected, not every subset', async () => {
    const fetchMock = mockFetch();
    const registry = new EmbeddedFontRegistry();
    await discoverFontsFromDocument(documentUsing('TestFont'), {}, undefined, registry);

    const fontFileCalls = fetchMock.mock.calls.filter(([url]) => url.endsWith('.ttf'));
    expect(fontFileCalls).toHaveLength(1);
    expect(fontFileCalls[0][0]).toContain('latin');
  });

  it('never fetches anything for a family the document does not actually use', async () => {
    const fetchMock = mockFetch();
    const registry = new EmbeddedFontRegistry();
    await discoverFontsFromDocument(documentUsing('SomeOtherFont'), {}, undefined, registry);

    expect(registry.has('TestFont')).toBe(false);
    expect(fetchMock.mock.calls.filter(([url]) => url.endsWith('.ttf'))).toHaveLength(0);
  });

  it('picks the Cyrillic subset instead when the document\'s own actual text is Cyrillic, not Latin', async () => {
    const fetchMock = mockFetch();
    const registry = new EmbeddedFontRegistry();
    const cyrillicDocument = {
      metadata: { fontFamily: 'TestFont', stylesheetLinks: ['https://example.com/fonts.css'] },
      children: [{ type: 'styledSpan', style: { fontFamily: 'TestFont' }, children: [{ type: 'text', value: 'АБВ' }] }], // АБВ
    };
    await discoverFontsFromDocument(cyrillicDocument, {}, undefined, registry);

    const fontFileCalls = fetchMock.mock.calls.filter(([url]) => url.endsWith('.ttf'));
    expect(fontFileCalls).toHaveLength(1);
    expect(fontFileCalls[0][0]).toContain('cyrillic');

    const entry = registry.match('TestFont', 400, false);
    expect(entry.sfnt.glyphIdForCodepoint(0x410)).toBeDefined(); // А
  });
});

describe('discoverFontsFromDocument — variable-font weight safety net', () => {
  afterEach(() => { vi.unstubAllGlobals(); clearStylesheetCaches(); });

  const VARIABLE_FONT = buildTestFont({ variable: true, glyphs: [{ cp: 0x41, width: 500 }] });

  const VARIABLE_CSS = `
    @font-face {
      font-family: 'VarFont';
      font-weight: 400;
      src: url(https://example.com/varfont.ttf);
    }
    @font-face {
      font-family: 'VarFont';
      font-weight: 700;
      src: url(https://example.com/varfont.ttf);
    }
  `;

  function mockFetch() {
    const fn = vi.fn(async (url) => ({
      ok: true,
      status: 200,
      statusText: 'OK',
      headers: new Map(),
      text: async () => VARIABLE_CSS,
      arrayBuffer: async () => VARIABLE_FONT.buffer.slice(VARIABLE_FONT.byteOffset, VARIABLE_FONT.byteOffset + VARIABLE_FONT.byteLength),
    }));
    vi.stubGlobal('fetch', fn);
    return fn;
  }

  function documentUsing(family) {
    return { metadata: { fontFamily: family, stylesheetLinks: ['https://example.com/fonts.css'] }, children: [] };
  }

  it('registers the regular (400, non-italic) weight from a variable-font source', async () => {
    mockFetch();
    const registry = new EmbeddedFontRegistry();
    await discoverFontsFromDocument(documentUsing('VarFont'), {}, undefined, registry);
    expect(registry.hasExact('VarFont', 400, false)).toBe(true);
  });

  it('refuses to register a non-regular weight from the same variable-font source, warning instead', async () => {
    mockFetch();
    const registry = new EmbeddedFontRegistry();
    const warnings = [];
    await discoverFontsFromDocument(documentUsing('VarFont'), {}, (code, msg) => warnings.push({ code, msg }), registry);

    expect(registry.hasExact('VarFont', 700, false)).toBe(false);
    expect(warnings.some((w) => w.code === 'VARIABLE_FONT_WEIGHT_SKIPPED')).toBe(true);
  });
});
