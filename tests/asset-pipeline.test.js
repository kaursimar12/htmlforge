/**
 * End-to-end checks on the shape of the pipeline itself (see CLAUDE.md's
 * Internal Architecture section): the asset resolver is the ONLY place that
 * ever fetches/reads image bytes, called exactly once per convert(), and
 * text exporters never trigger any I/O at all — not "resolveAssets happens
 * to dedupe," but "resolveAssets is only ever invoked once in the whole
 * pipeline for a given conversion."
 */
import { describe, it, expect, afterEach, afterAll, vi } from 'vitest';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { convert } from '../src/index.js';

const REMOTE_PNG_HTML =
  '<img src="https://example.com/logo.png"><img src="https://example.com/logo.png"><img src="https://example.com/logo.png">';

describe('asset pipeline shape', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  // A real 1x1 PNG — so PDF's own real byte decode succeeds cleanly too,
  // not just DOCX/PPTX/XLSX's "any bytes will do" media embedding.
  const TINY_PNG_BYTES = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    'base64',
  );

  function mockFetch() {
    const fn = vi.fn(async () => ({
      ok: true,
      status: 200,
      statusText: 'OK',
      headers: new Map(),
      arrayBuffer: async () => TINY_PNG_BYTES.buffer.slice(TINY_PNG_BYTES.byteOffset, TINY_PNG_BYTES.byteOffset + TINY_PNG_BYTES.byteLength),
    }));
    vi.stubGlobal('fetch', fn);
    return fn;
  }

  it('the same remote image referenced 3 times in one document is fetched exactly once', async () => {
    const fetchMock = mockFetch();
    await convert(REMOTE_PNG_HTML, { format: 'docx' });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  for (const format of ['markdown', 'txt', 'json', 'csv']) {
    it(`${format} never triggers a network fetch for a remote image`, async () => {
      const fetchMock = mockFetch();
      await convert(REMOTE_PNG_HTML, { format });
      expect(fetchMock).not.toHaveBeenCalled();
    });
  }

  it('every binary exporter (docx/pptx/pdf/xlsx) is flagged as needing resolved assets', async () => {
    const { getExporters } = await import('../src/exporter/registry.js');
    const flagged = getExporters().filter((e) => e.embedsAssets).map((e) => e.id).sort();
    expect(flagged).toEqual(['docx', 'pdf', 'pptx', 'xlsx']);
  });

  // Regression test: font loading/discovery was originally (wrongly) gated
  // on the SAME `embedsAssets` flag as images, which is also true for
  // docx/pptx/xlsx — meaning a DOCX/PPTX export where a document links a
  // stylesheet would have triggered real font file reads/network fetches
  // those exporters never even read from `ctx.fonts` (they only ever write a
  // plain font-family name string, see fonts/resolver.js's own
  // `resolveFontFamilyName`). Only PDF (the one format fixed/rasterized at
  // generation time) needs real font bytes, so it alone carries
  // `embedsFonts: true`, checked separately in src/index.js.
  it('only PDF is flagged as needing real embedded font bytes', async () => {
    const { getExporters } = await import('../src/exporter/registry.js');
    const flagged = getExporters().filter((e) => e.embedsFonts).map((e) => e.id).sort();
    expect(flagged).toEqual(['pdf']);
  });

  it('every binary exporter consumes ctx.assets rather than resolving its own', async () => {
    const fetchMock = mockFetch();
    for (const format of ['docx', 'pptx', 'pdf', 'xlsx']) {
      fetchMock.mockClear();
      const { warnings } = await convert('<img src="https://example.com/logo.png">', { format });
      expect(warnings).toEqual([]);
      expect(fetchMock).toHaveBeenCalledTimes(1);
    }
  });

  // convert()'s own input contract accepts a file PATH, not just an HTML
  // string (see src/index.js's resolveInput) — a relative <img src> in that
  // file means relative to THAT file's own directory (the same thing a
  // browser opening it locally would do), not wherever this process happens
  // to be running from. Real files on disk, not a data: URI, since this is
  // specifically exercising path resolution against the INPUT file's own
  // location.
  describe('relative image paths resolve against the input HTML FILE\'s own directory', () => {
    let dir;

    afterAll(async () => {
      if (dir) await rm(dir, { recursive: true, force: true });
    });

    it('a file-path input auto-derives baseUrl from its own directory, with no baseUrl option needed', async () => {
      dir = await mkdtemp(path.join(tmpdir(), 'html-parsor-asset-test-'));
      await writeFile(path.join(dir, 'logo.png'), TINY_PNG_BYTES);
      const htmlPath = path.join(dir, 'page.html');
      await writeFile(htmlPath, '<img src="./logo.png" alt="logo">');

      const { warnings } = await convert(htmlPath, { format: 'docx' });
      expect(warnings).toEqual([]);
    });

    it('an explicit options.baseUrl still overrides the derived one', async () => {
      const htmlPath = path.join(dir, 'page.html');
      const { warnings } = await convert(htmlPath, { format: 'docx', baseUrl: path.join(dir, 'does-not-exist') });
      expect(warnings.some((w) => w.code === 'MISSING_IMAGE')).toBe(true);
    });
  });
});
