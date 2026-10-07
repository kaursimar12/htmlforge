/**
 * PPTX exporter — IDM → PresentationML (.pptx).
 * Milestones: PP1 (skeleton) · PP2 (sections/headings → slides) ·
 *             PP3 (title/body placeholders) · PP4 (bullet lists) ·
 *             PP5 (images) · PP6 (tables). PP7 (speaker notes) is deferred —
 *             it would consume FootnoteNode, which the HTML parser doesn't
 *             produce yet (see project memory).
 */
import { ZipWriter } from '../../utils/zip-writer.js';
import { contentTypesXml, rootRelsXml, corePropsXml, appPropsXml, RelationshipManager } from '../ooxml/package.js';
import {
  resolveSlideSize, computeSlideGeometry, presentationXml, buildPresentationRels,
  themeXml, slideMasterXml, buildSlideMasterRels, slideLayoutXml, slideLayoutRelsXml,
  PPTX_REL_TYPES,
} from './layout.js';
import { collectSlideGroups, classifyGroup, renderSlidePage } from './slide.js';
import { resolveFontFamilyName } from '../../fonts/resolver.js';

const CONTENT_TYPE_OVERRIDES = [
  { partName: '/ppt/presentation.xml', contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml' },
  { partName: '/ppt/slideMasters/slideMaster1.xml', contentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml' },
  { partName: '/ppt/slideLayouts/slideLayout1.xml', contentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml' },
  { partName: '/ppt/theme/theme1.xml', contentType: 'application/vnd.openxmlformats-officedocument.theme+xml' },
];
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

export const pptxExporter = {
  id: 'pptx',
  extensions: ['.pptx'],
  mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  // A non-embedded image reference doesn't make sense for a self-contained
  // .pptx — convert() (src/index.js) reads this flag to resolve real image
  // bytes exactly once, upstream, before export() ever runs.
  embedsAssets: true,

  async export(document, ctx) {
    const slideSize = resolveSlideSize(ctx.options);
    const geometry = computeSlideGeometry(slideSize);
    const groups = collectSlideGroups(document);

    // ctx.assets was already resolved once by convert() itself — this
    // exporter only ever decides how to EMBED a resolved asset, never how
    // to obtain one (no fetch/readFile here or anywhere else in this file).
    const mediaFiles = [];
    const imageExtensions = [];
    const assetInfoByAssetId = new Map();
    for (const asset of ctx.assets ?? []) {
      if (!asset.data) continue; // resolver couldn't fetch bytes; already warned by convert()
      const filename = `image${mediaFiles.length + 1}.${asset.extension}`;
      mediaFiles.push({ filename, data: asset.data });
      imageExtensions.push(asset.extension);
      assetInfoByAssetId.set(asset.id, { filename, width: asset.width, height: asset.height });
    }

    // One GROUP can span several actual slides once its content doesn't fit
    // `geometry.bottomY` (see slide.js's renderSlidePage) — so the slide
    // count isn't known until after rendering, unlike the old one-group-one-
    // slide mapping. Only a group's own title (if any) is passed to its
    // first page; every continuation page renders with no title, picking up
    // exactly where the previous page's overflow left off.
    const slideParts = [];
    for (const group of groups) {
      const { title, units } = classifyGroup(group);
      let pageTitle = title;
      let pageUnits = units;
      do {
        const relationships = new RelationshipManager();
        relationships.add(PPTX_REL_TYPES.slideLayout, '../slideLayouts/slideLayout1.xml');
        // `body{font-size:...}` (dom-to-idm.js's own metadata.fontSizePt) is
        // what a plain, unstyled paragraph/list item/table cell should
        // actually render at — see shapes.js's own resolveDefaultFontPt,
        // which every relevant render/estimate function reads this through,
        // instead of always falling back to a flat 18pt regardless of what
        // the source document itself declared. Carried unscaled (a plain
        // CSS pt value) — resolveDefaultFontPt applies the same
        // CSS_TO_SLIDE_FONT_SCALE every other CSS-derived font size in that
        // file already goes through.
        const rctx = { relationships, assetInfoByAssetId, images: new Map(), bodyFontSizePt: document.metadata?.fontSizePt };
        const { xml, remainingUnits } = renderSlidePage(pageTitle, pageUnits, rctx, geometry);
        slideParts.push({ xml, relsXml: relationships.toXml() });
        pageTitle = undefined;
        pageUnits = remainingUnits;
      } while (pageUnits.length > 0);
    }

    const { xml: presentationRelsXmlStr, masterRId, slideRIds } = buildPresentationRels(slideParts.length);
    const { xml: slideMasterRelsXmlStr, layoutRId } = buildSlideMasterRels();

    const overrides = [
      ...CONTENT_TYPE_OVERRIDES,
      ...slideParts.map((_, i) => ({ partName: `/ppt/slides/slide${i + 1}.xml`, contentType: SLIDE_CONTENT_TYPE })),
    ];

    const zip = new ZipWriter();
    zip.addText('[Content_Types].xml', contentTypesXml({ imageExtensions, overrides }));
    zip.addText('_rels/.rels', rootRelsXml('ppt/presentation.xml'));
    zip.addText('docProps/core.xml', corePropsXml(document.metadata));
    zip.addText('docProps/app.xml', appPropsXml());
    zip.addText('ppt/presentation.xml', presentationXml(masterRId, slideRIds, slideSize));
    zip.addText('ppt/_rels/presentation.xml.rels', presentationRelsXmlStr);
    zip.addText('ppt/theme/theme1.xml', themeXml(document.metadata?.fontFamily ? resolveFontFamilyName(document.metadata.fontFamily) : undefined));
    zip.addText('ppt/slideMasters/slideMaster1.xml', slideMasterXml(slideSize, layoutRId));
    zip.addText('ppt/slideMasters/_rels/slideMaster1.xml.rels', slideMasterRelsXmlStr);
    zip.addText('ppt/slideLayouts/slideLayout1.xml', slideLayoutXml(slideSize));
    zip.addText('ppt/slideLayouts/_rels/slideLayout1.xml.rels', slideLayoutRelsXml());
    slideParts.forEach(({ xml, relsXml }, i) => {
      zip.addText(`ppt/slides/slide${i + 1}.xml`, xml);
      zip.addText(`ppt/slides/_rels/slide${i + 1}.xml.rels`, relsXml);
    });
    for (const { filename, data } of mediaFiles) {
      zip.addFile(`ppt/media/${filename}`, data);
    }

    return zip.finalize();
  },
};
