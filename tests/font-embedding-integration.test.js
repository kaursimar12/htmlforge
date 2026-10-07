/**
 * End-to-end checks that the font-resolution pipeline is actually wired
 * together correctly through convert() — not just that its individual
 * pieces (sfnt-parser/resolver/embedded-font) work in isolation (see
 * sfnt-parser.test.js/font-resolver.test.js/font-face-css.test.js), but that
 * a real bundled font (see fonts/builtin-substitutes.js) produces a real
 * embedded CID font in the output PDF, and that DOCX/PPTX still only ever
 * write a plain name (no embedding, no network) regardless.
 */
import { describe, it, expect } from 'vitest';
import { convert } from '../src/index.js';
import { ZipArchive } from '../src/utils/zip.js';

describe('font embedding, wired end-to-end through convert()', () => {
  it('PDF: a bundled font is embedded as a real Type0/CIDFontType2/FontFile2 chain', async () => {
    const html = '<p style="font-family:Poppins;font-weight:700">Hello</p>';
    const { buffer, warnings } = await convert(html, { format: 'pdf' });

    expect(warnings).toEqual([]);

    const text = buffer.toString('latin1');
    expect(text).toContain('/Subtype /Type0');
    expect(text).toContain('/CIDFontType2');
    expect(text).toContain('/Encoding /Identity-H');
    expect(text).toContain('/FontFile2');
    expect(text).toContain('/CIDToGIDMap /Identity');
    expect(text).toContain('/BaseFont /poppins-700');
  });

  it('PDF: a font-family with no registered source falls back to a Standard-14 font, never fails', async () => {
    const html = '<p style="font-family:SomeFontNotBundled">Hello</p>';
    const { warnings } = await convert(html, { format: 'pdf' });
    expect(warnings).toEqual([]);
  });

  it('DOCX: never embeds a bundled font\'s bytes at all — writes the plain family name unchanged', async () => {
    const html = '<p style="font-family:Poppins">Hello</p>';
    const { buffer, warnings } = await convert(html, { format: 'docx' });
    expect(warnings).toEqual([]);
    const docXml = ZipArchive.read(buffer).readText('word/document.xml');
    expect(docXml).toContain('w:ascii="Poppins"');
    expect(docXml).not.toContain('FontFile2');
  });

  it('DOCX: a bare CSS generic keyword resolves to a real installable font name', async () => {
    const html = '<p style="font-family:sans-serif">Hello</p>';
    const { buffer } = await convert(html, { format: 'docx' });
    const docXml = ZipArchive.read(buffer).readText('word/document.xml');
    expect(docXml).toContain('w:ascii="Arial"');
    expect(docXml).not.toContain('w:ascii="sans-serif"');
  });

  it('PPTX: a bare CSS generic keyword resolves to a real installable font name', async () => {
    const html = '<p style="font-family:monospace">Hello</p>';
    const { buffer } = await convert(html, { format: 'pptx' });
    const slideXml = ZipArchive.read(buffer).readText('ppt/slides/slide1.xml');
    expect(slideXml).toContain('typeface="Courier New"');
  });
});
