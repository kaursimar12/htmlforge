import { describe, it, expect, afterEach, vi } from 'vitest';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { resolveAssets } from '../src/assets/resolver.js';

function docWithImages(...srcs) {
  return {
    type: 'document',
    children: srcs.map((src) => ({ type: 'image', src })),
  };
}

describe('resolveAssets', () => {
  it('decodes a data: URI and stamps assetId', async () => {
    const pngBytes = Buffer.from([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
      0x00, 0x00, 0x00, 0x0d,
    ]);
    const dataUri = `data:image/png;base64,${pngBytes.toString('base64')}`;
    const doc = docWithImages(dataUri);

    const { assets, warnings } = await resolveAssets(doc);

    expect(warnings).toHaveLength(0);
    expect(assets).toHaveLength(1);
    expect(assets[0].data).toEqual(pngBytes);
    expect(assets[0].mimeType).toBe('image/png');
    expect(doc.children[0].assetId).toBe(assets[0].id);
  });

  it('a resolved asset carries real intrinsic width/height and a filename extension, computed once', async () => {
    // A real 1x1 PNG (same fixture DOCX's own tests use).
    const dataUri = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
    const { assets } = await resolveAssets(docWithImages(dataUri));

    expect(assets[0].width).toBe(1);
    expect(assets[0].height).toBe(1);
    expect(assets[0].extension).toBe('png');
  });

  it('an undecodable format (e.g. GIF) has no width/height, but still gets a filename extension', async () => {
    const dataUri = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7';
    const { assets } = await resolveAssets(docWithImages(dataUri));

    expect(assets[0].width).toBeUndefined();
    expect(assets[0].height).toBeUndefined();
    expect(assets[0].extension).toBe('gif');
  });

  it('a nonexistent relative local path → warning, no throw', async () => {
    const doc = docWithImages('./photo.png');
    const { assets, warnings } = await resolveAssets(doc);

    expect(warnings).toHaveLength(1);
    expect(warnings[0].code).toBe('MISSING_IMAGE');
    expect(assets[0].data).toBeUndefined();
  });

  it('deduplicates: same src → same assetId, one Asset record', async () => {
    const src = 'data:text/plain,hello';
    const doc = {
      type: 'document',
      children: [
        { type: 'image', src },
        { type: 'image', src },
      ],
    };

    const { assets } = await resolveAssets(doc);

    expect(assets).toHaveLength(1);
    const id0 = doc.children[0].assetId;
    const id1 = doc.children[1].assetId;
    expect(id0).toBe(id1);
  });

  it('embedImages: false — skips fetching, asset has no data', async () => {
    const doc = docWithImages('data:image/png;base64,abc');
    const { assets } = await resolveAssets(doc, { embedImages: false });

    expect(assets[0].data).toBeUndefined();
    expect(assets[0].src).toBe('data:image/png;base64,abc');
  });

  it('missing local file → warning, no throw', async () => {
    const doc = docWithImages('/no/such/file.png');
    const { assets, warnings } = await resolveAssets(doc);

    expect(warnings).toHaveLength(1);
    expect(warnings[0].code).toBe('MISSING_IMAGE');
    expect(assets[0].data).toBeUndefined();
  });

  it('strictAssets: true + missing file → throws', async () => {
    const doc = docWithImages('/no/such/file.png');
    await expect(resolveAssets(doc, { strictAssets: true })).rejects.toThrow('strictAssets');
  });

  it('guesses MIME type from file extension when magic bytes unknown', async () => {
    const doc2 = docWithImages('data:image/jpeg;base64,/9j/AA==');
    const { assets } = await resolveAssets(doc2);
    expect(assets[0].mimeType).toBe('image/jpeg');
  });

  it('walks images nested inside sections and lists', async () => {
    const doc = {
      type: 'document',
      children: [
        {
          type: 'section',
          children: [
            { type: 'image', src: 'data:text/plain,a' },
          ],
        },
        {
          type: 'list',
          ordered: false,
          items: [
            {
              type: 'listItem',
              children: [{ type: 'image', src: 'data:text/plain,b' }],
            },
          ],
        },
      ],
    };

    const { assets } = await resolveAssets(doc);
    expect(assets).toHaveLength(2);
  });

  it('walks an image that is a direct child of a table cell (no strong/emphasis/link wrapper)', async () => {
    const doc = {
      type: 'document',
      children: [
        {
          type: 'table',
          rows: [
            { type: 'tableRow', cells: [{ type: 'tableCell', children: [{ type: 'image', src: 'data:text/plain,a' }] }] },
          ],
        },
      ],
    };

    const { assets } = await resolveAssets(doc);
    expect(assets).toHaveLength(1);
  });

  it('walks an image nested in a paragraph inside a table cell (block-content cell)', async () => {
    const doc = {
      type: 'document',
      children: [
        {
          type: 'table',
          rows: [
            {
              type: 'tableRow',
              cells: [{
                type: 'tableCell',
                children: [{ type: 'paragraph', children: [{ type: 'image', src: 'data:text/plain,a' }] }],
              }],
            },
          ],
        },
      ],
    };

    const { assets } = await resolveAssets(doc);
    expect(assets).toHaveLength(1);
  });

  it('walks an image wrapped in a styledSpan', async () => {
    const doc = {
      type: 'document',
      children: [
        { type: 'paragraph', children: [{ type: 'styledSpan', style: { bold: true }, children: [{ type: 'image', src: 'data:text/plain,a' }] }] },
      ],
    };

    const { assets } = await resolveAssets(doc);
    expect(assets).toHaveLength(1);
  });
});

describe('resolveAssets — local files', () => {
  let dir;

  afterEach(async () => {
    if (dir) await rm(dir, { recursive: true, force: true });
    dir = undefined;
  });

  it('reads a real local file relative to baseUrl (as a directory)', async () => {
    dir = await mkdtemp(path.join(tmpdir(), 'html-parsor-'));
    const bytes = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await writeFile(path.join(dir, 'photo.png'), bytes);

    const doc = docWithImages('./photo.png');
    const { assets, warnings } = await resolveAssets(doc, { baseUrl: dir });

    expect(warnings).toHaveLength(0);
    expect(assets[0].data).toEqual(bytes);
    expect(assets[0].mimeType).toBe('image/png');
  });

  it('with no baseUrl, a relative path resolves against process.cwd()', async () => {
    dir = await mkdtemp(path.join(tmpdir(), 'html-parsor-'));
    const bytes = Buffer.from('plain', 'utf8');
    await writeFile(path.join(dir, 'note.png'), bytes);

    const relative = path.relative(process.cwd(), path.join(dir, 'note.png'));
    const doc = docWithImages(relative);
    const { assets, warnings } = await resolveAssets(doc);

    expect(warnings).toHaveLength(0);
    expect(assets[0].data).toEqual(bytes);
  });
});

describe('resolveAssets — http(s) fetch', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function mockFetch(impl) {
    vi.stubGlobal('fetch', vi.fn(impl));
  }

  it('fetches a remote image and stamps real bytes', async () => {
    const bytes = new Uint8Array([1, 2, 3, 4]);
    mockFetch(async () => ({
      ok: true,
      status: 200,
      statusText: 'OK',
      headers: new Map(),
      arrayBuffer: async () => bytes.buffer,
    }));

    const doc = docWithImages('http://example.com/photo.png');
    const { assets, warnings } = await resolveAssets(doc);

    expect(warnings).toHaveLength(0);
    expect(assets[0].data).toEqual(Buffer.from(bytes));
  });

  it('a non-2xx response → warning, no throw', async () => {
    mockFetch(async () => ({ ok: false, status: 404, statusText: 'Not Found', headers: new Map() }));

    const doc = docWithImages('https://example.com/missing.png');
    const { assets, warnings } = await resolveAssets(doc);

    expect(warnings).toHaveLength(1);
    expect(warnings[0].code).toBe('MISSING_IMAGE');
    expect(warnings[0].message).toContain('404');
    expect(assets[0].data).toBeUndefined();
  });

  it('a response declaring Content-Length over maxImageBytes is rejected before reading the body', async () => {
    const arrayBuffer = vi.fn();
    mockFetch(async () => ({
      ok: true,
      status: 200,
      statusText: 'OK',
      headers: new Map([['content-length', '999999']]),
      arrayBuffer,
    }));

    const doc = docWithImages('https://example.com/huge.png');
    const { warnings } = await resolveAssets(doc, { maxImageBytes: 10 });

    expect(warnings).toHaveLength(1);
    expect(warnings[0].message).toContain('maxImageBytes');
    expect(arrayBuffer).not.toHaveBeenCalled();
  });

  it('a body larger than maxImageBytes is rejected even with no Content-Length header', async () => {
    const bytes = new Uint8Array(100);
    mockFetch(async () => ({
      ok: true,
      status: 200,
      statusText: 'OK',
      headers: new Map(),
      arrayBuffer: async () => bytes.buffer,
    }));

    const doc = docWithImages('https://example.com/photo.png');
    const { warnings, assets } = await resolveAssets(doc, { maxImageBytes: 10 });

    expect(warnings).toHaveLength(1);
    expect(assets[0].data).toBeUndefined();
  });

  it('an aborted (timed out) fetch → a clear timeout warning, no throw', async () => {
    mockFetch(async () => {
      const err = new Error('aborted');
      err.name = 'AbortError';
      throw err;
    });

    const doc = docWithImages('https://example.com/slow.png');
    const { warnings } = await resolveAssets(doc, { imageTimeoutMs: 5 });

    expect(warnings).toHaveLength(1);
    expect(warnings[0].message).toContain('timed out');
  });

  it('baseUrl as a URL resolves a relative src into an absolute URL before fetching', async () => {
    const bytes = new Uint8Array([9, 9]);
    let requestedUrl;
    mockFetch(async (url) => {
      requestedUrl = url;
      return { ok: true, status: 200, statusText: 'OK', headers: new Map(), arrayBuffer: async () => bytes.buffer };
    });

    const doc = docWithImages('images/chart.png');
    const { assets } = await resolveAssets(doc, { baseUrl: 'https://myapp.com/reports/' });

    expect(String(requestedUrl)).toBe('https://myapp.com/reports/images/chart.png');
    expect(assets[0].data).toEqual(Buffer.from(bytes));
  });
});
