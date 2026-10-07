/**
 * PDF exporter — IDM → a from-scratch PDF file (no npm dependencies).
 * Milestones: PDF1 (skeleton) · PDF2 (text/paragraph flow) · PDF3 (headings)
 *   · PDF4 (inline formatting) · PDF5 (lists) · PDF6 (page breaks, auto +
 *   explicit) · PDF7 (images) · PDF8 (tables) · PDF9 (hyperlinks) ·
 *   PDF10 (metadata + page setup)
 */
import { resolvePageGeometry } from './layout.js';
import { PdfPageFlow } from './page-flow.js';
import { renderBlockSequence } from './blocks.js';
import { buildImageXObject } from './images.js';
import { buildPdf } from './pdf-writer.js';

export const pdfExporter = {
  id: 'pdf',
  extensions: ['.pdf'],
  mimeType: 'application/pdf',
  // A non-embedded image reference doesn't make sense for a self-contained
  // .pdf — convert() (src/index.js) reads this flag to resolve real image
  // bytes exactly once, upstream, before export() ever runs.
  embedsAssets: true,
  // PDF is the only format fixed/rasterized at generation time, so it's the
  // only one that needs real font BYTES (see fonts/resolver.js's own doc
  // comment) — DOCX/PPTX/XLSX just write a font-family name string and
  // trust the reading machine's own installed fonts. convert() gates the
  // font-loading/network-discovery work specifically on THIS flag, not on
  // embedsAssets (which is also true for docx/pptx/xlsx), so those formats
  // never trigger a font file fetch/parse they'd never even read.
  embedsFonts: true,

  async export(document, ctx) {
    // ctx.assets was already resolved once by convert() itself — this
    // exporter only ever decides how to EMBED a resolved asset (as a real
    // DCTDecode/FlateDecode XObject — see images.js), never how to obtain
    // one.
    const imageXObjectByAssetId = new Map();
    let imageCounter = 0;
    for (const asset of ctx.assets ?? []) {
      const xObject = buildImageXObject(asset, `Im${imageCounter + 1}`, (code, message) => ctx.warn(code, message));
      if (xObject) {
        imageCounter++;
        imageXObjectByAssetId.set(asset.id, xObject);
      }
    }

    const geometry = resolvePageGeometry(ctx.options, document.metadata?.contentWidthPx, document.metadata?.pageMarginPt);
    // The page's own declared base font-size/line-height (see dom-to-idm.js's
    // metadata.fontSizePt/lineHeight) — this project's own default only
    // kicks in when the HTML never declared one (see PdfPageFlow's own
    // constructor).
    const flow = new PdfPageFlow(geometry, {
      bodyFontPt: document.metadata?.fontSizePt,
      lineHeightMultiplier: document.metadata?.lineHeight,
      // ctx.fonts is resolved once by convert() (src/index.js), the same
      // "resolved upstream, exporter only decides how to embed/use it"
      // relationship ctx.assets already has to the image resolver.
      embeddedFontRegistry: ctx.fonts,
      warn: (code, message) => ctx.warn(code, message),
    });
    renderBlockSequence(flow, document.children ?? [], imageXObjectByAssetId);

    if (flow.hasUnsupportedGlyphs) {
      ctx.warn('pdf-unsupported-glyph', 'One or more characters have no WinAnsiEncoding representation in the standard PDF fonts (e.g. CJK, Arabic, Cyrillic, emoji), or no glyph in an embedded font, and were replaced with "?"/a missing-glyph box.');
    }

    const { pages, fontEntries, imageEntries, embeddedFontEntries } = flow.finish();
    return buildPdf({ pages, fontEntries, imageEntries, embeddedFontEntries, metadata: document.metadata ?? {} });
  },
};
