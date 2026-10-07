/**
 * Low-level PDF file format writer — object numbering, indirect-object
 * serialization, classic table `xref`, `trailer`, `startxref`/`%%EOF`. Built
 * on `ByteWriter` (utils/bytes.js), mirroring `ZipWriter`'s own three-phase
 * shape (body written sequentially while offsets are recorded → a second
 * section built from those offsets → a trailer, then `Buffer.concat`).
 *
 * Has no opinion on page layout — `page-flow.js` produces the finished,
 * in-memory `{pages, fontEntries, imageEntries, metadata}` description this
 * module consumes, with no object numbers assigned yet. Numbers are only
 * ever allocated here, once, in one fixed pass — page count, font usage,
 * and image usage are all only knowable after the full document walk
 * completes (the same reason `ZipWriter` doesn't know its central
 * directory contents until every `addFile` call is done), so nothing here
 * ever needs renumbering once allocation starts.
 */
import { deflateSync } from 'node:zlib';
import { ByteWriter } from '../../utils/bytes.js';
import { pdfString } from './content-stream.js';

export class Ref {
  constructor(num) { this.num = num; }
}
export const ref = (num) => new Ref(num);

class Name {
  constructor(value) { this.value = value; }
}
export const name = (value) => new Name(value);

function formatNumber(n) {
  if (Number.isInteger(n)) return String(n);
  return String(Math.round(n * 1000) / 1000);
}

// A minimal, generic PDF value serializer — dicts/arrays/refs/names/strings/
// numbers/booleans. Every object this exporter writes (Catalog, Pages,
// Page, Font, XObject, Annot, Info) is built as a plain JS object/array
// passed through this, rather than hand-formatting each one's syntax
// separately — one correct serializer is less error-prone than several
// hand-written ones for the same handful of value kinds. `undefined` values
// in a plain object are dropped (the "omit this key entirely" convention
// used throughout, e.g. `/Title` when there's no title).
export function serialize(value) {
  if (value instanceof Ref) return `${value.num} 0 R`;
  if (value instanceof Name) return `/${value.value}`;
  if (Array.isArray(value)) return `[${value.map(serialize).join(' ')}]`;
  if (typeof value === 'number') return formatNumber(value);
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  if (typeof value === 'string') return `(${pdfString(value)})`;
  if (value && typeof value === 'object') {
    const parts = [];
    for (const [key, v] of Object.entries(value)) {
      if (v === undefined) continue;
      parts.push(`/${key} ${serialize(v)}`);
    }
    return `<<${parts.map((p) => ' ' + p).join('')} >>`;
  }
  throw new Error(`Cannot serialize PDF value: ${JSON.stringify(value)}`);
}

// One xref entry is exactly 20 bytes — the classic table format requires
// this exact fixed width per spec (10-digit offset, space, 5-digit
// generation, space, 'n'/'f', space, EOL).
function xrefEntry(offset) {
  return `${String(offset).padStart(10, '0')} 00000 n \n`;
}

class PdfObjectWriter {
  #body = new ByteWriter();
  #offsetByObjNum = new Map();

  constructor() {
    // Written into #body itself (not concatenated on afterward) so every
    // object's own recorded offset is already correct in final-file
    // coordinates the moment writeObject computes it — concatenating the
    // header on separately at the end would silently shift every offset in
    // the xref table by the header's own byte length. The four bytes after
    // the version comment (each ≥0x80) are PDF's own conventional "binary
    // file" signal for tools that sniff the first few bytes of a file
    // transferred in text mode.
    this.#body.writeString('%PDF-1.4\n');
    this.#body.writeBytes(Buffer.from([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));
  }

  writeObject(objNum, dictValue, streamBuffer) {
    this.#offsetByObjNum.set(objNum, this.#body.length);
    this.#body.writeString(`${objNum} 0 obj\n`);
    if (streamBuffer) {
      this.#body.writeString(serialize({ ...dictValue, Length: streamBuffer.length }));
      this.#body.writeString('\nstream\n');
      this.#body.writeBytes(streamBuffer);
      this.#body.writeString('\nendstream');
    } else {
      this.#body.writeString(serialize(dictValue));
    }
    this.#body.writeString('\nendobj\n');
  }

  finalize(rootRef, infoRef) {
    const bodyBuffer = this.#body.toBuffer();
    const xrefOffset = bodyBuffer.length;
    const maxObjNum = Math.max(0, ...this.#offsetByObjNum.keys());

    const xref = new ByteWriter();
    xref.writeString(`xref\n0 ${maxObjNum + 1}\n`);
    xref.writeString('0000000000 65535 f \n');
    for (let n = 1; n <= maxObjNum; n++) {
      xref.writeString(xrefEntry(this.#offsetByObjNum.get(n) ?? 0));
    }

    const trailerDict = { Size: maxObjNum + 1, Root: rootRef, Info: infoRef };
    const trailer = new ByteWriter();
    trailer.writeString('trailer\n');
    trailer.writeString(serialize(trailerDict));
    trailer.writeString(`\nstartxref\n${xrefOffset}\n%%EOF`);

    return Buffer.concat([bodyBuffer, xref.toBuffer(), trailer.toBuffer()]);
  }
}

const PRODUCER = 'html-parsor PDF exporter';

/**
 * Allocates every object number in one fixed pass (Catalog, Pages root,
 * fonts, images, then per page a page dict + content stream + one object
 * per link annotation, then an optional Info last) and writes the whole
 * file. Nothing here is pre-numbered by the caller — `page-flow.js` hands
 * over plain, number-free descriptions:
 *
 * - `fontEntries`: `[{ resourceName, baseFont }]` (from `FontRegistry.entries()`)
 * - `imageEntries`: `[{ resourceName, dict, data }]` — `dict` is the
 *   XObject dict minus `/Length` (added automatically from `data.length`)
 * - `pages`: `[{ width, height, ops, usedImageResourceNames: Set<string>,
 *   links: [{ rect: [x0,y0,x1,y1], uri }] }]`
 * - `metadata`: `{ title, author, language }`
 */
export function buildPdf({ pages, fontEntries, imageEntries, embeddedFontEntries = [], metadata }) {
  const writer = new PdfObjectWriter();

  const catalogObjNum = 1;
  const pagesObjNum = 2;
  let nextObjNum = 3;

  const fontObjNumByResourceName = new Map();
  for (const font of fontEntries) {
    fontObjNumByResourceName.set(font.resourceName, nextObjNum++);
  }

  // Each embedded font is a CHAIN of 5 objects (Type0 -> [CIDFontType2] ->
  // FontDescriptor -> FontFile2, plus a sibling ToUnicode stream off the
  // Type0 dict) rather than the single object a Standard-14 font needs, or
  // the single optional SMask sibling an image needs — see embedded-font.js's
  // own doc comment for why. Every number is allocated here, up front, same
  // "describe with placeholders now, resolve into refs once numbers exist"
  // two-pass shape images.js's SMask handling already established.
  const embeddedFontObjNums = new Map();
  for (const font of embeddedFontEntries) {
    embeddedFontObjNums.set(font.resourceName, {
      type0: nextObjNum++,
      cidFont: nextObjNum++,
      fontDescriptor: nextObjNum++,
      fontFile: nextObjNum++,
      toUnicode: nextObjNum++,
    });
  }

  // A PNG that carried an alpha channel (images.js's own pngXObject) hands
  // back its alpha as a separate sibling XObject (`image.smask`) rather
  // than baking it into the color image's own dict directly — object
  // numbers don't exist yet at that point (this is the one place in the
  // whole exporter that allocates them, per this module's own doc comment),
  // so it can only carry the smask's *description*, not a real `/SMask`
  // reference. Each sibling gets its own object number allocated here,
  // right next to its own parent's.
  const imageObjNumByResourceName = new Map();
  for (const image of imageEntries) {
    imageObjNumByResourceName.set(image.resourceName, nextObjNum++);
    if (image.smask) imageObjNumByResourceName.set(image.smask.resourceName, nextObjNum++);
  }

  const pageAllocations = pages.map((page) => {
    const objNum = nextObjNum++;
    const contentObjNum = nextObjNum++;
    const linkObjNums = page.links.map(() => nextObjNum++);
    return { page, objNum, contentObjNum, linkObjNums };
  });

  const hasInfo = Boolean(metadata.title || metadata.author);
  const infoObjNum = hasInfo ? nextObjNum++ : undefined;

  for (const font of fontEntries) {
    writer.writeObject(fontObjNumByResourceName.get(font.resourceName), {
      Type: name('Font'),
      Subtype: name('Type1'),
      BaseFont: name(font.baseFont),
      Encoding: name('WinAnsiEncoding'),
    });
  }

  for (const image of imageEntries) {
    // `image.dict.SMask` is `true` (a placeholder, see images.js's own
    // pngXObject) rather than a real reference until now, since only this
    // module ever knows the sibling's actual object number.
    const dict = image.smask ? { ...image.dict, SMask: ref(imageObjNumByResourceName.get(image.smask.resourceName)) } : image.dict;
    writer.writeObject(imageObjNumByResourceName.get(image.resourceName), dict, image.data);
    if (image.smask) writer.writeObject(imageObjNumByResourceName.get(image.smask.resourceName), image.smask.dict, image.smask.data);
  }

  for (const font of embeddedFontEntries) {
    const nums = embeddedFontObjNums.get(font.resourceName);
    const { descendantFont, toUnicode } = font;
    const { fontDescriptor } = descendantFont;
    const { fontFile } = fontDescriptor;

    writer.writeObject(nums.fontFile, fontFile.dict, fontFile.data);
    writer.writeObject(nums.fontDescriptor, { ...fontDescriptor.dict, FontFile2: ref(nums.fontFile) });
    writer.writeObject(nums.cidFont, { ...descendantFont.dict, FontDescriptor: ref(nums.fontDescriptor) });
    writer.writeObject(nums.toUnicode, toUnicode.dict, toUnicode.data);
    writer.writeObject(nums.type0, {
      ...font.dict,
      DescendantFonts: [ref(nums.cidFont)],
      ToUnicode: ref(nums.toUnicode),
    });
  }

  // Standard fonts have always been listed on every page's Resources.Font
  // regardless of whether that specific page draws any text at all (unlike
  // images, which are filtered per-page via usedImageResourceNames) — an
  // embedded font's own weight is entirely in its single shared FontFile2
  // object, not duplicated per page reference, so listing it on every page
  // too costs nothing extra and keeps one consistent convention for every
  // kind of font this exporter can produce.
  const fontResourceDict = {};
  for (const font of fontEntries) fontResourceDict[font.resourceName] = ref(fontObjNumByResourceName.get(font.resourceName));
  for (const font of embeddedFontEntries) fontResourceDict[font.resourceName] = ref(embeddedFontObjNums.get(font.resourceName).type0);

  const pageRefs = [];
  for (const { page, objNum, contentObjNum, linkObjNums } of pageAllocations) {
    const xObjectResourceDict = {};
    for (const resourceName of page.usedImageResourceNames) {
      xObjectResourceDict[resourceName] = ref(imageObjNumByResourceName.get(resourceName));
    }

    const contentStream = deflateSync(Buffer.from(page.ops, 'latin1'));
    writer.writeObject(contentObjNum, { Filter: name('FlateDecode') }, contentStream);

    const annotRefs = page.links.map((link, i) => {
      writer.writeObject(linkObjNums[i], {
        Type: name('Annot'),
        Subtype: name('Link'),
        Rect: link.rect,
        Border: [0, 0, 0],
        A: { S: name('URI'), URI: link.uri },
      });
      return ref(linkObjNums[i]);
    });

    writer.writeObject(objNum, {
      Type: name('Page'),
      Parent: ref(pagesObjNum),
      MediaBox: [0, 0, page.width, page.height],
      Resources: {
        Font: fontResourceDict,
        XObject: Object.keys(xObjectResourceDict).length ? xObjectResourceDict : undefined,
      },
      Contents: ref(contentObjNum),
      Annots: annotRefs.length ? annotRefs : undefined,
    });
    pageRefs.push(ref(objNum));
  }

  writer.writeObject(pagesObjNum, { Type: name('Pages'), Kids: pageRefs, Count: pageRefs.length });
  writer.writeObject(catalogObjNum, {
    Type: name('Catalog'),
    Pages: ref(pagesObjNum),
    Lang: metadata.language || undefined,
  });

  let infoRef;
  if (hasInfo) {
    writer.writeObject(infoObjNum, {
      Title: metadata.title || undefined,
      Author: metadata.author || undefined,
      Producer: PRODUCER,
    });
    infoRef = ref(infoObjNum);
  }

  return writer.finalize(ref(catalogObjNum), infoRef);
}
