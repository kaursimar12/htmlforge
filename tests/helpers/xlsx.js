/**
 * Unzips a generated .xlsx and parses every XML part, for structural
 * assertions in tests — not just "does the ZIP start with PK." Mirrors
 * tests/helpers/docx.js.
 */

import { ZipArchive } from '../../src/utils/zip.js';
import { parseXml } from '../../src/utils/xml.js';

export function readXlsxParts(buffer) {
  const zip = ZipArchive.read(buffer);
  const xml = {};
  for (const name of zip.entries()) {
    if (name.endsWith('.xml') || name.endsWith('.rels')) {
      xml[name] = parseXml(zip.readText(name));
    }
  }
  return { zip, xml };
}
