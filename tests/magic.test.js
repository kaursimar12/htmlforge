import { describe, expect, it } from 'vitest';
import { detectFormat } from '../src/utils/magic.js';
import { makeZip } from './helpers/zip.js';

describe('detectFormat', () => {
  it('detects binary signatures', () => {
    expect(detectFormat(Buffer.from('%PDF-1.7\n'))?.format).toBe('pdf');
    expect(detectFormat(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))?.format).toBe('png');
    expect(detectFormat(Buffer.from([0xff, 0xd8, 0xff, 0xe0]))?.format).toBe('jpeg');
    expect(detectFormat(Buffer.from('GIF89a'))?.format).toBe('gif');
  });

  it('disambiguates the OOXML family by archive contents', () => {
    expect(detectFormat(makeZip({ 'word/document.xml': '<w:document/>' }))?.format).toBe('docx');
    expect(detectFormat(makeZip({ 'xl/workbook.xml': '<workbook/>' }))?.format).toBe('xlsx');
    expect(detectFormat(makeZip({ 'ppt/presentation.xml': '<p:presentation/>' }))?.format).toBe('pptx');
  });

  it('falls back to generic zip for unknown archives', () => {
    expect(detectFormat(makeZip({ 'random.bin': 'x' }))?.format).toBe('zip');
  });

  it('sniffs html and xml text', () => {
    expect(detectFormat(Buffer.from('<!DOCTYPE html><html></html>'))?.format).toBe('html');
    expect(detectFormat(Buffer.from('<?xml version="1.0"?><root/>'))?.format).toBe('xml');
  });

  it('returns undefined for plain text', () => {
    expect(detectFormat(Buffer.from('just some words'))).toBeUndefined();
  });
});
