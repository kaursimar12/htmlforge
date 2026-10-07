import { describe, expect, it } from 'vitest';
import { ZipArchive, isZip } from '../src/utils/zip.js';
import { makeZip } from './helpers/zip.js';

describe('ZipArchive', () => {
  it('reads deflated members', () => {
    const buf = makeZip({ 'word/document.xml': '<w:document/>', 'a/b.txt': 'hello' });
    const zip = ZipArchive.read(buf);
    expect(zip.entries()).toEqual(['word/document.xml', 'a/b.txt']);
    expect(zip.has('word/document.xml')).toBe(true);
    expect(zip.readText('word/document.xml')).toBe('<w:document/>');
    expect(zip.readText('a/b.txt')).toBe('hello');
  });

  it('reads stored (uncompressed) members', () => {
    const buf = makeZip({ 'note.txt': 'stored content' }, { deflate: false });
    expect(ZipArchive.read(buf).readText('note.txt')).toBe('stored content');
  });

  it('returns undefined for absent members and lists entries by prefix', () => {
    const buf = makeZip({ 'ppt/media/image1.png': 'x', 'ppt/slides/slide1.xml': 'y' });
    const zip = ZipArchive.read(buf);
    expect(zip.read('missing')).toBeUndefined();
    expect(zip.entriesUnder('ppt/media/')).toEqual(['ppt/media/image1.png']);
  });

  it('round-trips a larger payload through deflate', () => {
    const big = 'lorem ipsum '.repeat(5000);
    const zip = ZipArchive.read(makeZip({ 'big.txt': big }));
    expect(zip.readText('big.txt')).toBe(big);
  });

  it('rejects non-zip input', () => {
    expect(() => ZipArchive.read(Buffer.from('not a zip'))).toThrow(/zip/i);
  });
});

describe('isZip', () => {
  it('detects the local-file-header signature', () => {
    expect(isZip(makeZip({ 'a.txt': 'a' }))).toBe(true);
    expect(isZip(Buffer.from('%PDF-1.7'))).toBe(false);
  });
});
