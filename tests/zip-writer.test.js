import { describe, expect, it } from 'vitest';
import { ZipWriter } from '../src/utils/zip-writer.js';
import { ZipArchive } from '../src/utils/zip.js';

describe('ZipWriter', () => {
  it('round-trips text files through ZipArchive.read()', () => {
    const zip = new ZipWriter();
    zip.addText('hello.txt', 'Hello, World!');
    zip.addText('nested/file.xml', '<root/>');
    const buf = zip.finalize();

    const archive = ZipArchive.read(buf);
    expect(archive.readText('hello.txt')).toBe('Hello, World!');
    expect(archive.readText('nested/file.xml')).toBe('<root/>');
  });

  it('lists all entries', () => {
    const zip = new ZipWriter();
    zip.addText('a.txt', 'a').addText('b.txt', 'b');
    const buf = zip.finalize();
    const archive = ZipArchive.read(buf);
    expect(archive.entries()).toContain('a.txt');
    expect(archive.entries()).toContain('b.txt');
  });

  it('round-trips binary data (stored mode)', () => {
    const data = Buffer.from([0x00, 0x01, 0x02, 0xff]);
    const zip = new ZipWriter();
    zip.addFile('bin.dat', data, false);
    const archive = ZipArchive.read(zip.finalize());
    expect(archive.read('bin.dat')).toEqual(data);
  });

  it('returns undefined for missing entry', () => {
    const zip = new ZipWriter();
    zip.addText('a.txt', 'a');
    const archive = ZipArchive.read(zip.finalize());
    expect(archive.read('missing.txt')).toBeUndefined();
  });

  it('produces a buffer starting with PK magic bytes', () => {
    const buf = new ZipWriter().addText('x', 'y').finalize();
    expect(buf[0]).toBe(0x50);
    expect(buf[1]).toBe(0x4b);
  });
});
