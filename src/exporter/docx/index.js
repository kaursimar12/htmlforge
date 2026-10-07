/**
 * DOCX exporter — IDM → Office Open XML WordprocessingML (.docx).
 * Milestones: W1 (skeleton) · W2 (paragraphs) · W3 (headings/breaks) ·
 *             W4 (inline formatting) · W5 (lists) · W6 (tables) ·
 *             W7 (hyperlinks) · W8 (images) · W9 (metadata) · W10 (page setup)
 */

import { ZipWriter } from '../../utils/zip-writer.js';
import { contentTypesXml, rootRelsXml, corePropsXml, appPropsXml, RelationshipManager, REL_TYPES } from '../ooxml/package.js';
import { stylesXml } from './styles.js';
import { NumberingManager } from './numbering.js';
import { buildDocumentXml } from './document.js';

export const docxExporter = {
  id: 'docx',
  extensions: ['.docx'],
  mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  // A non-embedded image reference doesn't make sense for a self-contained
  // .docx — convert() (src/index.js) reads this flag to resolve real image
  // bytes exactly once, upstream, before export() ever runs.
  embedsAssets: true,

  async export(document, ctx) {
    const relationships = new RelationshipManager();
    relationships.add(REL_TYPES.styles, 'styles.xml');
    relationships.add(REL_TYPES.numbering, 'numbering.xml');

    // ctx.assets was already resolved once by convert() itself — this
    // exporter only ever decides how to EMBED a resolved asset, never how
    // to obtain one (no fetch/readFile here or anywhere else in this file).
    const mediaFiles = [];
    const imageExtensions = [];
    const imagesByAssetId = new Map();
    for (const asset of ctx.assets ?? []) {
      if (!asset.data) continue; // resolver couldn't fetch bytes; already warned by convert()
      const filename = `image${mediaFiles.length + 1}.${asset.extension}`;
      mediaFiles.push({ filename, data: asset.data });
      imageExtensions.push(asset.extension);
      const relId = relationships.add(REL_TYPES.image, `media/${filename}`);
      imagesByAssetId.set(asset.id, { relId, width: asset.width, height: asset.height });
    }

    const numbering = new NumberingManager();
    let imageCounter = 0;
    let bookmarkCounter = 0;
    const rctx = {
      relationships,
      numbering,
      images: imagesByAssetId,
      nextImageId: () => ++imageCounter,
      nextBookmarkId: () => bookmarkCounter++,
    };

    // Rendering populates `numbering` as a side effect (each <ul>/<ol> found
    // registers itself), so document.xml must be built before numbering.xml.
    const documentXml = buildDocumentXml(document, ctx, rctx);

    const zip = new ZipWriter();
    zip.addText('[Content_Types].xml', contentTypesXml({ imageExtensions }));
    zip.addText('_rels/.rels', rootRelsXml());
    zip.addText('docProps/core.xml', corePropsXml(document.metadata));
    zip.addText('docProps/app.xml', appPropsXml());
    zip.addText('word/document.xml', documentXml);
    zip.addText('word/styles.xml', stylesXml(document.metadata));
    zip.addText('word/numbering.xml', numbering.toXml());
    zip.addText('word/_rels/document.xml.rels', relationships.toXml());
    for (const { filename, data } of mediaFiles) {
      zip.addFile(`word/media/${filename}`, data);
    }

    return zip.finalize();
  },
};
