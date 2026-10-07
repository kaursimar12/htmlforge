/**
 * Build a minimal ZIP archive in-memory for tests, independent of ZipWriter.
 */
import { deflateRawSync } from 'node:zlib';

export function makeZip(files, options = {}) {
  const deflate = options.deflate ?? true;
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const [name, content] of Object.entries(files)) {
    const nameBuf = Buffer.from(name, 'utf8');
    const raw = Buffer.from(content, 'utf8');
    const method = deflate ? 8 : 0;
    const data = deflate ? deflateRawSync(raw) : raw;

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt32LE(0, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    const localEntry = Buffer.concat([local, nameBuf, data]);
    localParts.push(localEntry);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0, 8);
    cd.writeUInt16LE(method, 10);
    cd.writeUInt32LE(0, 16);
    cd.writeUInt32LE(data.length, 20);
    cd.writeUInt32LE(raw.length, 24);
    cd.writeUInt16LE(nameBuf.length, 28);
    cd.writeUInt32LE(offset, 42);
    centralParts.push(Buffer.concat([cd, nameBuf]));

    offset += localEntry.length;
  }

  const localBuf = Buffer.concat(localParts);
  const centralBuf = Buffer.concat(centralParts);

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(centralParts.length, 8);
  eocd.writeUInt16LE(centralParts.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(localBuf.length, 16);

  return Buffer.concat([localBuf, centralBuf, eocd]);
}
