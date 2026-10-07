/**
 * Coverage for the opt-in Google Fonts fallback — only ever tried for a
 * family the document uses but nothing else (discovered @font-face, bundled
 * fonts) already resolved. Same fetch-mocking pattern as
 * stylesheet-resolver.test.js, since this module reuses its CSS-parsing/
 * decode machinery wholesale.
 */
import { describe, it, expect, afterEach, vi } from 'vitest';
import { EmbeddedFontRegistry } from '../src/fonts/resolver.js';
import { fetchGoogleFontsFallback, clearGoogleFontsCache } from '../src/fonts/google-fonts-fallback.js';
import { buildTestFont } from './helpers/sfnt.js';

describe('fetchGoogleFontsFallback', () => {
  afterEach(() => { vi.unstubAllGlobals(); clearGoogleFontsCache(); });

  const FONT_400 = buildTestFont({ glyphs: [{ cp: 0x41, width: 500 }] });
  const FONT_700 = buildTestFont({ glyphs: [{ cp: 0x41, width: 600 }] });

  const CSS = `
    @font-face {
      font-family: 'Poppins';
      font-weight: 400;
      src: url(https://fonts.gstatic.com/poppins-400.woff2);
    }
    @font-face {
      font-family: 'Poppins';
      font-weight: 700;
      src: url(https://fonts.gstatic.com/poppins-700.woff2);
    }
  `;

  function mockFetch({ cssOk = true } = {}) {
    const fn = vi.fn(async (url) => {
      if (String(url).includes('fonts.googleapis.com')) {
        if (!cssOk) return { ok: false, status: 404, statusText: 'Not Found' };
        return { ok: true, status: 200, statusText: 'OK', text: async () => CSS };
      }
      const body = String(url).includes('700') ? FONT_700 : FONT_400;
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        headers: new Map(),
        arrayBuffer: async () => body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength),
      };
    });
    vi.stubGlobal('fetch', fn);
    return fn;
  }

  function documentUsing(family) {
    return { metadata: { fontFamily: family }, children: [] };
  }

  it('resolves a family with no other source via the Google Fonts CSS2 API', async () => {
    mockFetch();
    const registry = new EmbeddedFontRegistry();
    await fetchGoogleFontsFallback(documentUsing('Poppins'), {}, undefined, registry);

    expect(registry.hasExact('Poppins', 400, false)).toBe(true);
    expect(registry.hasExact('Poppins', 700, false)).toBe(true);
  });

  it('never runs for a family already registered by an earlier source', async () => {
    const fetchMock = mockFetch();
    const registry = new EmbeddedFontRegistry();
    const font = buildTestFont();
    registry.register('Poppins', 400, false, undefined, font);

    await fetchGoogleFontsFallback(documentUsing('Poppins'), {}, undefined, registry);

    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('warns and leaves the family unresolved when Google Fonts does not have it', async () => {
    mockFetch({ cssOk: false });
    const registry = new EmbeddedFontRegistry();
    const warnings = [];
    await fetchGoogleFontsFallback(documentUsing('NotARealFont'), {}, (code, msg) => warnings.push({ code, msg }), registry);

    expect(registry.has('NotARealFont')).toBe(false);
    expect(warnings).toHaveLength(1);
    expect(warnings[0].code).toBe('GOOGLE_FONT_LOOKUP_FAILED');
  });

  it('never fetches anything when every used family is already resolved', async () => {
    const fetchMock = mockFetch();
    const registry = new EmbeddedFontRegistry();
    registry.register('Arial', 400, false, undefined, buildTestFont());
    await fetchGoogleFontsFallback(documentUsing('Arial'), {}, undefined, registry);

    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('retries with Title Case when the CSS-declared casing fails, and resolves on the second try', async () => {
    const fn = vi.fn(async (url) => {
      const urlStr = String(url);
      if (urlStr.includes('fonts.googleapis.com')) {
        if (urlStr.includes('family=LATO')) return { ok: false, status: 400, statusText: 'Bad Request' };
        if (urlStr.includes('family=Lato')) return { ok: true, status: 200, statusText: 'OK', text: async () => CSS.replace(/Poppins/g, 'Lato') };
        throw new Error(`unexpected URL: ${urlStr}`);
      }
      const body = urlStr.includes('700') ? FONT_700 : FONT_400;
      return { ok: true, status: 200, statusText: 'OK', headers: new Map(), arrayBuffer: async () => body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength) };
    });
    vi.stubGlobal('fetch', fn);

    const registry = new EmbeddedFontRegistry();
    await fetchGoogleFontsFallback(documentUsing('LATO'), {}, undefined, registry);

    expect(registry.hasExact('Lato', 400, false)).toBe(true);
    expect(fn.mock.calls.filter(([url]) => String(url).includes('fonts.googleapis.com'))).toHaveLength(2);
  });

  it('never retries with Title Case when the CSS-declared casing already succeeds', async () => {
    const fetchMock = mockFetch();
    const registry = new EmbeddedFontRegistry();
    await fetchGoogleFontsFallback(documentUsing('Poppins'), {}, undefined, registry);

    expect(fetchMock.mock.calls.filter(([url]) => String(url).includes('fonts.googleapis.com'))).toHaveLength(1);
  });

  it('refuses to register a non-regular weight resolved to a variable font, warning instead', async () => {
    const VARIABLE_FONT = buildTestFont({ variable: true, glyphs: [{ cp: 0x41, width: 500 }] });
    const VARIABLE_CSS = `
      @font-face { font-family: 'VarFont'; font-weight: 400; src: url(https://fonts.gstatic.com/varfont.ttf); }
      @font-face { font-family: 'VarFont'; font-weight: 700; src: url(https://fonts.gstatic.com/varfont.ttf); }
    `;
    const fn = vi.fn(async (url) => {
      const urlStr = String(url);
      if (urlStr.includes('fonts.googleapis.com')) return { ok: true, status: 200, statusText: 'OK', text: async () => VARIABLE_CSS };
      return { ok: true, status: 200, statusText: 'OK', headers: new Map(), arrayBuffer: async () => VARIABLE_FONT.buffer.slice(VARIABLE_FONT.byteOffset, VARIABLE_FONT.byteOffset + VARIABLE_FONT.byteLength) };
    });
    vi.stubGlobal('fetch', fn);

    const registry = new EmbeddedFontRegistry();
    const warnings = [];
    await fetchGoogleFontsFallback(documentUsing('VarFont'), {}, (code, msg) => warnings.push({ code, msg }), registry);

    expect(registry.hasExact('VarFont', 400, false)).toBe(true);
    expect(registry.hasExact('VarFont', 700, false)).toBe(false);
    expect(warnings.some((w) => w.code === 'VARIABLE_FONT_WEIGHT_SKIPPED')).toBe(true);
  });
});
