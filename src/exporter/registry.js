/**
 * Exporter registry — selection by format id or file extension.
 */

import { markdownExporter } from './markdown/index.js';
import { csvExporter } from './csv/index.js';
import { txtExporter } from './txt/index.js';
import { jsonExporter } from './json/index.js';
import { docxExporter } from './docx/index.js';
import { xlsxExporter } from './xlsx/index.js';
import { pptxExporter } from './pptx/index.js';
import { pdfExporter } from './pdf/index.js';

const exporters = [markdownExporter, csvExporter, txtExporter, jsonExporter, docxExporter, xlsxExporter, pptxExporter, pdfExporter];

export function registerExporter(exporter) {
  const existing = exporters.findIndex((e) => e.id === exporter.id);
  if (existing >= 0) exporters[existing] = exporter;
  else exporters.unshift(exporter);
}

export function getExporters() {
  return exporters;
}

export function getExporterById(id) {
  return exporters.find((e) => e.id === id);
}

export function getExporterByExtension(ext) {
  const normalized = ext.toLowerCase();
  return exporters.find((e) => e.extensions.includes(normalized));
}

export function selectExporter(format) {
  const byId = getExporterById(format);
  if (byId) return byId;

  const ext = format.startsWith('.') ? format : `.${format}`;
  const byExt = getExporterByExtension(ext);
  if (byExt) return byExt;

  const available = exporters.map((e) => e.id).join(', ') || 'none registered yet';
  throw new Error(`Unsupported format "${format}". Available: ${available}`);
}
