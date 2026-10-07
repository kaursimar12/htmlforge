/**
 * html-exporter — Universal HTML → Document exporter.
 *
 * HTML is the only input. Any supported format is the output.
 * One API. One input format. Any output.
 *
 * @example
 * ```js
 * import { convert } from 'html-exporter';
 *
 * const result = await convert('<h1>Hello</h1><p>World</p>', { format: 'docx' });
 * await writeFile('output.docx', result.buffer);
 * ```
 */

import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { parseHtml } from './html-parser/index.js';
import { normalizeDocument } from './normalizer/index.js';
import { resolveScrollableContent } from './idm/resolve-scrollable-content.js';
import { selectExporter } from './exporter/registry.js';
import { createContext } from './exporter/types.js';
import { resolveAssets } from './assets/resolver.js';
import { EmbeddedFontRegistry } from './fonts/resolver.js';
import { discoverFontsFromDocument } from './fonts/stylesheet-resolver.js';
import { loadBuiltinSubstituteFonts } from './fonts/builtin-substitutes.js';
import { fetchGoogleFontsFallback } from './fonts/google-fonts-fallback.js';

const DEFAULT_OPTIONS = {
  embedImages: false,
  strictAssets: false,
  preserveScrollableContent: true,
};

/**
 * Convert HTML to any supported document format.
 *
 * @param input   HTML string, file path, or Buffer/Uint8Array of HTML bytes.
 * @param options Format selection and exporter configuration.
 */
export async function convert(input, options = {}) {
  const format = options.format ?? options.exporter;
  if (!format) {
    throw new Error('options.format is required. Example: { format: "docx" }');
  }

  const { html, baseUrl: inferredBaseUrl } = await resolveInput(input);

  // A file-path input's own folder is the default root for its relative
  // image srcs — the same rule a browser already applies when it resolves
  // <img src="./images/x.png"> against the page's own URL. An explicit
  // options.baseUrl always wins; this only fills the gap when the caller
  // didn't set one, and stays undefined (falling through to the resolver's
  // own process.cwd() fallback — see assets/resolver.js) for raw HTML
  // strings/Buffers, which have no file location to infer from.
  const resolved = { ...DEFAULT_OPTIONS, baseUrl: inferredBaseUrl, ...options };

  const warnings = [];
  const assets = [];
  const fonts = new EmbeddedFontRegistry();

  const exporter = selectExporter(format);
  const ctx = createContext(resolved, warnings, assets, fonts);

  const idm = parseHtml(html);
  const resolvedScrollable = resolveScrollableContent(idm, resolved, ctx.warn);
  const normalized = normalizeDocument(resolvedScrollable);

  // Inject convert options into metadata
  if (options.title || options.author || options.language) {
    normalized.metadata = {
      ...normalized.metadata,
      title: options.title ?? normalized.metadata?.title,
      author: options.author ?? normalized.metadata?.author,
      language: options.language ?? normalized.metadata?.language,
    };
  }

  // The asset resolver is the ONLY place that ever fetches/reads image
  // bytes — called exactly once here, before the exporter runs, so every
  // format consumes the identical resolved assets (see ctx.assets) instead
  // of each independently re-fetching the same images. Only the binary,
  // self-contained formats (`exporter.embedsAssets`) need real bytes at
  // all — Markdown/TXT/JSON/CSV only ever print a resolved *reference*
  // (utils/html.js's displayImageSrc, via the same pure path-resolution
  // logic the resolver itself uses) and must never trigger a fetch or a
  // file read.
  if (exporter.embedsAssets) {
    // A non-embedded image reference doesn't make sense for a self-contained
    // binary document — embedImages is forced on here regardless of
    // convert()'s own default.
    const { assets: resolvedAssets, warnings: assetWarnings } = await resolveAssets(normalized, { ...resolved, embedImages: true });
    assets.push(...resolvedAssets);
    warnings.push(...assetWarnings);
  }

  // Real font BYTES only ever matter to a format fixed at generation time
  // (PDF, today — see pdf/index.js's own embedsFonts flag) — DOCX/PPTX/XLSX
  // just write a font-family name string and trust the reading machine's
  // own installed fonts, so they never read `ctx.fonts` at all and must
  // never trigger a font file fetch/parse for nothing. Deliberately gated
  // on `embedsFonts`, NOT `embedsAssets` (which is also true for
  // docx/pptx/xlsx) — conflating the two would make a DOCX/PPTX export with
  // a document linking a stylesheet do real network I/O it never needed and
  // nothing ever reads.
  if (exporter.embedsFonts) {
    // Discovery from the page's OWN declared stylesheets/@font-face —
    // unconditional, not gated behind a flag: unlike the Google Fonts
    // fallback below (which GUESSES a URL with nothing in the document
    // backing it), every stylesheet/font URL this fetches is one the HTML
    // itself already named — following a reference the source explicitly
    // gave, not making a request on the caller's behalf to somewhere the
    // document never pointed. `fonts` (the EmbeddedFontRegistry instance)
    // was already handed to createContext above — this fills it in place,
    // the same live-reference trick `assets`/`warnings` already use above.
    await discoverFontsFromDocument(normalized, resolved, (code, message) => ctx.warn(code, message), fonts);

    // Bundled OFL substitutes/fonts (see fonts/builtin-substitutes.js) —
    // local files only, unconditional like the step above; never overrides
    // a combo already resolved by it.
    await loadBuiltinSubstituteFonts(normalized, fonts, (code, message) => ctx.warn(code, message));

    // Opt-in (default false), deliberately provider-specific exception to
    // every other font-resolution step here — see google-fonts-fallback.js's
    // own doc comment for why this is kept separate from the provider-
    // agnostic discoverFontsFromDocument above. Runs last: only ever
    // attempted for families the steps above left completely unresolved.
    if (resolved.googleFontsFallback) {
      await fetchGoogleFontsFallback(normalized, resolved, (code, message) => ctx.warn(code, message), fonts);
    }
  }

  const buffer = await exporter.export(normalized, ctx);

  const metadata = {
    format: exporter.id,
    ...normalized.metadata,
  };

  return {
    buffer,
    mimeType: exporter.mimeType,
    extension: exporter.extensions[0],
    metadata,
    warnings,
  };
}

async function resolveInput(input) {
  if (typeof input === 'string') {
    // HTML string starts with '<', or is whitespace then '<'; an empty
    // string isn't a valid file path either, so per the input contract's
    // "otherwise treated as an HTML string" fallback it's an empty document,
    // with no file location to infer a baseUrl from.
    if (input === '' || input.trimStart().startsWith('<')) return { html: input };
    // Otherwise treat as a file path — its own directory becomes the
    // default root for any relative image src inside it.
    const fileBytes = await readFile(input);
    return { html: fileBytes.toString('utf8'), baseUrl: path.dirname(path.resolve(input)) };
  }
  const buf = Buffer.isBuffer(input) ? input : Buffer.from(input);
  return { html: buf.toString('utf8') };
}

/* -------------------------------------------------------------------------- */
/* Re-exports for advanced usage                                               */
/* -------------------------------------------------------------------------- */

export { registerExporter, getExporters } from './exporter/registry.js';
export { parseHtml } from './html-parser/index.js';
export * as idm from './idm/index.js';
