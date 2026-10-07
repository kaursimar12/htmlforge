/**
 * Content-based format detection from magic numbers.
 */

import { startsWith } from './bytes.js';
import { ZipArchive, isZip } from './zip.js';

const SIGNATURES = [
  { sig: [0x25, 0x50, 0x44, 0x46, 0x2d], format: 'pdf', mimeType: 'application/pdf', extension: '.pdf' },
  { sig: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], format: 'png', mimeType: 'image/png', extension: '.png' },
  { sig: [0xff, 0xd8, 0xff], format: 'jpeg', mimeType: 'image/jpeg', extension: '.jpg' },
  { sig: [0x47, 0x49, 0x46, 0x38], format: 'gif', mimeType: 'image/gif', extension: '.gif' },
  { sig: [0x42, 0x4d], format: 'bmp', mimeType: 'image/bmp', extension: '.bmp' },
  { sig: [0x49, 0x49, 0x2a, 0x00], format: 'tiff', mimeType: 'image/tiff', extension: '.tiff' },
  { sig: [0x4d, 0x4d, 0x00, 0x2a], format: 'tiff', mimeType: 'image/tiff', extension: '.tiff' },
  { sig: [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], format: 'ole', mimeType: 'application/x-ole-storage', extension: '.bin' },
];

export function detectFormat(buffer) {
  for (const entry of SIGNATURES) {
    if (startsWith(buffer, entry.sig))
      return { format: entry.format, mimeType: entry.mimeType, extension: entry.extension };
  }

  if (startsWith(buffer, [0x52, 0x49, 0x46, 0x46]) && buffer.length >= 12 && buffer.toString('latin1', 8, 12) === 'WEBP')
    return { format: 'webp', mimeType: 'image/webp', extension: '.webp' };

  if (isZip(buffer))
    return detectZip(buffer);

  return detectText(buffer);
}

function detectZip(buffer) {
  try {
    const zip = ZipArchive.read(buffer);
    if (zip.has('word/document.xml'))
      return { format: 'docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', extension: '.docx' };
    if (zip.has('xl/workbook.xml'))
      return { format: 'xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', extension: '.xlsx' };
    if (zip.has('ppt/presentation.xml'))
      return { format: 'pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', extension: '.pptx' };
  } catch { /* fall through */ }
  return { format: 'zip', mimeType: 'application/zip', extension: '.zip' };
}

function detectText(buffer) {
  const head = buffer.subarray(0, 512).toString('utf8').replace(/^﻿/, '').trimStart();
  const lower = head.toLowerCase();
  if (lower.startsWith('<!doctype html') || lower.startsWith('<html') || lower.includes('<head'))
    return { format: 'html', mimeType: 'text/html', extension: '.html' };
  if (head.startsWith('<?xml') || (head.startsWith('<') && !lower.startsWith('<!doctype')))
    return { format: 'xml', mimeType: 'application/xml', extension: '.xml' };
  return undefined;
}
