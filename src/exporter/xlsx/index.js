/**
 * XLSX exporter — IDM → Office Open XML SpreadsheetML (.xlsx).
 * Milestones: X1 (skeleton) · X2 (tables→sheets, A1 refs, value
 * classification) · X3 (cell styling) · X4 (shared strings) · X5 (sheet
 * naming) · X6 (column auto-width) · X7 (merged cells) — plus later
 * additions: rich-text runs, vertical align, direction/writing-mode/
 * rotation, row height, and rendering the WHOLE document (not just tables)
 * into one sheet, stacked in document order (document-sheet.js).
 */

import { ZipWriter } from '../../utils/zip-writer.js';
import { contentTypesXml, rootRelsXml, corePropsXml, appPropsXml, RelationshipManager, REL_TYPES } from '../ooxml/package.js';
import { SharedStringsManager } from './shared-strings.js';
import { CellStyleManager } from './styles.js';
import { documentSheetXml } from './document-sheet.js';
import { resolveDocumentSheetName, buildWorkbookRels, workbookXml } from './workbook.js';
import { collectDocumentFlow } from '../../idm/collect-document-flow.js';
import { drawingXml, drawingRelsXml } from './drawing.js';

const CONTENT_TYPE_OVERRIDES = [
  { partName: '/xl/workbook.xml', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' },
  { partName: '/xl/styles.xml', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' },
  { partName: '/xl/sharedStrings.xml', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml' },
];

const WORKSHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';

export const xlsxExporter = {
  id: 'xlsx',
  extensions: ['.xlsx'],
  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  // A non-embedded image reference doesn't make sense for a self-contained
  // .xlsx — convert() (src/index.js) reads this flag to resolve real image
  // bytes exactly once, upstream, before export() ever runs.
  embedsAssets: true,

  async export(document, ctx) {
    const blocks = collectDocumentFlow(document.children);

    // A zero-worksheet .xlsx isn't a legal file (unlike CSV's "empty buffer
    // is a legal empty CSV") — Excel can't open it at all, so this warns and
    // still emits one valid, empty sheet rather than erroring.
    if (blocks.length === 0) {
      ctx.warn('no-content', 'No exportable content found in the document; XLSX output has one empty sheet');
    }

    const tables = blocks.filter((b) => b.type === 'table');
    const sheetName = resolveDocumentSheetName(document, tables, ctx.options);
    const { xml: workbookRelsXmlStr, sheetRelIds } = buildWorkbookRels(1);

    // X8 — ctx.assets was already resolved once by convert() itself; this
    // exporter only ever decides how to EMBED a resolved asset, never how to
    // obtain one (no fetch/readFile here or anywhere else in this file).
    // Resolved (data present), but only actually embedded below if
    // collectDocumentFlow's own block list places it — an image nested
    // inside a table cell or list item has no XLSX anchor representation
    // yet (see collect-document-flow.js), so it's resolved but never
    // referenced by an anchor; `dataByFilename` lets the zip step below
    // include only the media files an anchor actually points at.
    const dataByFilename = new Map();
    const imagesByAssetId = new Map();
    for (const asset of ctx.assets ?? []) {
      if (!asset.data) continue; // resolver couldn't fetch bytes; already warned by convert()
      const filename = `image${dataByFilename.size + 1}.${asset.extension}`;
      dataByFilename.set(filename, asset.data);
      imagesByAssetId.set(asset.id, { relId: `rId${dataByFilename.size}`, filename, width: asset.width, height: asset.height });
    }

    const sharedStrings = new SharedStringsManager();
    const styles = new CellStyleManager(document.metadata);
    // Carried alongside `styles` so document-sheet.js/worksheet.js's own row-
    // height estimates (which run before a cell's `s="N"` style is even
    // registered) size a plain, unstyled block/cell against the document's
    // own declared body font-size instead of this project's hardcoded
    // Calibri-11pt assumption — see those files' own DEFAULT_FONT_SIZE_PT.
    const defaultFontSizePt = document.metadata?.fontSizePt;
    const rctx = { sharedStrings, styles, defaultFontSizePt, images: imagesByAssetId };
    const { xml: sheetXmlBody, anchors } = documentSheetXml(blocks, rctx, document.metadata?.contentWidthPx);

    const workbookXmlStr = workbookXml([{ name: sheetName, rId: sheetRelIds[0] }]);

    const overrides = [
      ...CONTENT_TYPE_OVERRIDES,
      { partName: '/xl/worksheets/sheet1.xml', contentType: WORKSHEET_CONTENT_TYPE },
    ];

    // Only a document with at least one successfully embedded image gets a
    // drawing part at all — the common (image-free) case stays exactly the
    // ZIP layout it always was.
    let sheetXml = sheetXmlBody;
    let worksheetRelsXmlStr;
    const imageExtensions = anchors.map((a) => a.filename.split('.').pop());
    if (anchors.length > 0) {
      const worksheetRels = new RelationshipManager();
      const drawingRelId = worksheetRels.add(REL_TYPES.drawing, '../drawings/drawing1.xml');
      worksheetRelsXmlStr = worksheetRels.toXml();
      sheetXml = sheetXml.replace('</worksheet>', `<drawing r:id="${drawingRelId}"/></worksheet>`);
      overrides.push({ partName: '/xl/drawings/drawing1.xml', contentType: DRAWING_CONTENT_TYPE });
    }

    const zip = new ZipWriter();
    zip.addText('[Content_Types].xml', contentTypesXml({ imageExtensions, overrides }));
    zip.addText('_rels/.rels', rootRelsXml('xl/workbook.xml'));
    zip.addText('docProps/core.xml', corePropsXml(document.metadata));
    zip.addText('docProps/app.xml', appPropsXml());
    zip.addText('xl/workbook.xml', workbookXmlStr);
    zip.addText('xl/worksheets/sheet1.xml', sheetXml);
    zip.addText('xl/sharedStrings.xml', sharedStrings.toXml());
    zip.addText('xl/styles.xml', styles.toXml());
    zip.addText('xl/_rels/workbook.xml.rels', workbookRelsXmlStr);
    if (anchors.length > 0) {
      zip.addText('xl/worksheets/_rels/sheet1.xml.rels', worksheetRelsXmlStr);
      zip.addText('xl/drawings/drawing1.xml', drawingXml(anchors));
      zip.addText('xl/drawings/_rels/drawing1.xml.rels', drawingRelsXml(anchors));
      for (const anchor of anchors) {
        zip.addFile(`xl/media/${anchor.filename}`, dataByFilename.get(anchor.filename));
      }
    }

    return zip.finalize();
  },
};
