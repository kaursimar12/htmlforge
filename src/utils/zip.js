/**
 * Read-only ZIP reader built on Node's core zlib.
 */

import { inflateRawSync } from 'node:zlib';
import { ByteReader } from './bytes.js';
import { SIG_LOCAL, SIG_CENTRAL, SIG_EOCD, METHOD_STORED, METHOD_DEFLATE, FLAG_UTF8_NAME } from './zip-format.js';

export class ZipArchive {
  index = new Map();
  cache = new Map();

  constructor(buffer) {
    this.buffer = buffer;
  }

  static read(buffer) {
    const archive = new ZipArchive(buffer);
    archive.readCentralDirectory();
    return archive;
  }

  has(name) { return this.index.has(name); }

  entries() { return [...this.index.keys()]; }

  entriesUnder(prefix) {
    return this.entries().filter((name) => name.startsWith(prefix));
  }

  read(name) {
    const cached = this.cache.get(name);
    if (cached) return cached;
    const entry = this.index.get(name);
    if (!entry) return undefined;
    if (entry.flags & 0x1) throw new Error(`zip entry "${name}" is encrypted`);
    const data = this.extract(entry);
    this.cache.set(name, data);
    return data;
  }

  readText(name, encoding = 'utf8') {
    return this.read(name)?.toString(encoding);
  }

  extract(entry) {
    const r = new ByteReader(this.buffer).seek(entry.localHeaderOffset);
    if (r.u32le() !== SIG_LOCAL) throw new Error(`bad local header for "${entry.name}"`);
    r.skip(2 + 2 + 2 + 2 + 2 + 4 + 4 + 4);
    const nameLen = r.u16le();
    const extraLen = r.u16le();
    r.skip(nameLen + extraLen);
    const compressed = r.bytes(entry.compressedSize);
    if (entry.method === METHOD_STORED) return Buffer.from(compressed);
    if (entry.method === METHOD_DEFLATE) return inflateRawSync(compressed);
    throw new Error(`unsupported compression method ${entry.method} for "${entry.name}"`);
  }

  readCentralDirectory() {
    const eocd = this.findEocd();
    const r = new ByteReader(this.buffer).seek(eocd);
    r.skip(4);
    r.skip(2 + 2 + 2 + 2);
    r.skip(4);
    const cdOffset = r.u32le();
    if (cdOffset === 0xffffffff) throw new Error('Zip64 archives are not supported');
    let pos = cdOffset;
    while (pos + 4 <= this.buffer.length && this.buffer.readUInt32LE(pos) === SIG_CENTRAL) {
      pos = this.readCentralEntry(pos);
    }
  }

  readCentralEntry(offset) {
    const r = new ByteReader(this.buffer).seek(offset);
    r.skip(4);
    r.skip(2 + 2);
    const flags = r.u16le();
    const method = r.u16le();
    r.skip(2 + 2);
    r.skip(4);
    const compressedSize = r.u32le();
    const uncompressedSize = r.u32le();
    const nameLen = r.u16le();
    const extraLen = r.u16le();
    const commentLen = r.u16le();
    r.skip(2 + 2 + 4);
    const localHeaderOffset = r.u32le();
    const name = r.string(nameLen, flags & FLAG_UTF8_NAME ? 'utf8' : 'latin1');
    r.skip(extraLen + commentLen);
    if (!name.endsWith('/')) {
      this.index.set(name, { name, method, compressedSize, uncompressedSize, localHeaderOffset, flags });
    }
    return r.offset;
  }

  findEocd() {
    const buf = this.buffer;
    const minSize = 22;
    if (buf.length < minSize) throw new Error('not a zip archive (too small)');
    const scanStart = Math.max(0, buf.length - (minSize + 0xffff));
    for (let p = buf.length - minSize; p >= scanStart; p--) {
      if (buf.readUInt32LE(p) === SIG_EOCD) return p;
    }
    throw new Error('not a zip archive (no end-of-central-directory record)');
  }
}

export function isZip(buffer) {
  return buffer.length >= 4 && buffer.readUInt32LE(0) === SIG_LOCAL;
}
