/**
 * Shared OOXML package boilerplate — used by DOCX and XLSX (and, later,
 * PPTX, which is also a "ZIP of XML parts" the same way the other two are).
 */

import { XmlBuilder, xmlDoc } from '../../utils/xml-builder.js';

export const NS = {
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  packageRels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  officeRels: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  coreProps: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  extProps: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
};

export const REL_TYPES = {
  officeDocument: `${NS.officeRels}/officeDocument`,
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties: `${NS.officeRels}/extended-properties`,
  styles: `${NS.officeRels}/styles`,
  numbering: `${NS.officeRels}/numbering`,
  image: `${NS.officeRels}/image`,
  hyperlink: `${NS.officeRels}/hyperlink`,
  drawing: `${NS.officeRels}/drawing`,
  // SpreadsheetML-specific (xl/workbook.xml's own relationships).
  worksheet: `${NS.officeRels}/worksheet`,
  sharedStrings: `${NS.officeRels}/sharedStrings`,
};

const IMAGE_CONTENT_TYPES = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
};

// ─── [Content_Types].xml ────────────────────────────────────────────────────
// The docProps Overrides are generic across every OOXML package format; the
// main-part Overrides differ (WordprocessingML vs SpreadsheetML vs...), so
// `overrides` lets a caller other than docx supply its own list instead of
// docx's hardcoded `/word/*` ones. Omitted entirely (docx's existing calls)
// keeps today's behavior exactly as-is.

const DEFAULT_MAIN_PART_OVERRIDES = [
  { partName: '/word/document.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' },
  { partName: '/word/styles.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' },
  { partName: '/word/numbering.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml' },
];

export function contentTypesXml({ imageExtensions = [], overrides } = {}) {
  const b = new XmlBuilder();
  b.el('Types', { xmlns: NS.contentTypes });
  b.self('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' });
  b.self('Default', { Extension: 'xml', ContentType: 'application/xml' });
  for (const ext of new Set(imageExtensions)) {
    b.self('Default', { Extension: ext, ContentType: IMAGE_CONTENT_TYPES[ext] ?? 'application/octet-stream' });
  }
  for (const { partName, contentType } of overrides ?? DEFAULT_MAIN_PART_OVERRIDES) {
    b.self('Override', { PartName: partName, ContentType: contentType });
  }
  b.self('Override', {
    PartName: '/docProps/core.xml',
    ContentType: 'application/vnd.openxmlformats-package.core-properties+xml',
  });
  b.self('Override', {
    PartName: '/docProps/app.xml',
    ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  });
  return xmlDoc(b.end());
}

// ─── _rels/.rels ────────────────────────────────────────────────────────────
// `mainPart` is the package's own entry document — `word/document.xml` for
// DOCX, `xl/workbook.xml` for XLSX — everything else here (docProps) is
// identical across formats.

export function rootRelsXml(mainPart = 'word/document.xml') {
  const rels = new RelationshipManager();
  rels.add(REL_TYPES.officeDocument, mainPart);
  rels.add(REL_TYPES.coreProperties, 'docProps/core.xml');
  rels.add(REL_TYPES.extendedProperties, 'docProps/app.xml');
  return rels.toXml();
}

// ─── docProps/core.xml ──────────────────────────────────────────────────────

export function corePropsXml(metadata = {}) {
  const b = new XmlBuilder();
  b.el('cp:coreProperties', {
    'xmlns:cp': NS.coreProps,
    'xmlns:dc': NS.dc,
    'xmlns:dcterms': NS.dcterms,
    'xmlns:xsi': NS.xsi,
  });
  if (metadata.title) { b.el('dc:title').text(metadata.title).up(); }
  if (metadata.author) { b.el('dc:creator').text(metadata.author).up(); }
  if (metadata.language) { b.el('dc:language').text(metadata.language).up(); }
  if (metadata.createdAt) {
    const iso = metadata.createdAt instanceof Date ? metadata.createdAt.toISOString() : metadata.createdAt;
    b.el('dcterms:created', { 'xsi:type': 'dcterms:W3CDTF' }).text(iso).up();
  }
  return xmlDoc(b.end());
}

// ─── docProps/app.xml ───────────────────────────────────────────────────────

export function appPropsXml() {
  const b = new XmlBuilder();
  b.el('Properties', { xmlns: NS.extProps });
  b.el('Application').text('html-parsor').up();
  return xmlDoc(b.end());
}

// ─── Relationship management ───────────────────────────────────────────────
// One place that allocates rIds and builds a part's .rels file, so document.xml
// and its .rels file can never drift out of sync (the exact class of bug that
// makes Word show a "repair" prompt).

export class RelationshipManager {
  #rels = [];
  #counter = 1;

  add(type, target, targetMode) {
    const id = `rId${this.#counter++}`;
    this.#rels.push({ id, type, target, targetMode });
    return id;
  }

  toXml() {
    const b = new XmlBuilder();
    b.el('Relationships', { xmlns: NS.packageRels });
    for (const r of this.#rels) {
      b.self('Relationship', { Id: r.id, Type: r.type, Target: r.target, TargetMode: r.targetMode });
    }
    return xmlDoc(b.end());
  }
}
