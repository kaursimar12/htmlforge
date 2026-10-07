/**
 * ZIP archive writer built on Node's core zlib.
 */

import { deflateRawSync } from 'node:zlib';
import { ByteWriter } from './bytes.js';
import { SIG_LOCAL, SIG_CENTRAL, SIG_EOCD, METHOD_STORED, METHOD_DEFLATE, FLAG_UTF8_NAME } from './zip-format.js';

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export class ZipWriter {
  entries = [];
  bodyWriter = new ByteWriter();

  addFile(name, data, compress = true) {
    const localOffset = this.bodyWriter.length;
    const compressed = compress ? deflateRawSync(data) : data;

    const useCompressed = compress && compressed.length < data.length;
    const finalData = useCompressed ? compressed : data;
    const finalMethod = useCompressed ? METHOD_DEFLATE : METHOD_STORED;

    const nameBuf = Buffer.from(name, 'utf8');
    const crc = crc32(data);

    this.bodyWriter
      .writeU32le(SIG_LOCAL)
      .writeU16le(20)
      .writeU16le(FLAG_UTF8_NAME)
      .writeU16le(finalMethod)
      .writeU16le(0)
      .writeU16le(0)
      .writeU32le(crc)
      .writeU32le(finalData.length)
      .writeU32le(data.length)
      .writeU16le(nameBuf.length)
      .writeU16le(0)
      .writeBytes(nameBuf)
      .writeBytes(finalData);

    this.entries.push({
      name,
      data,
      compressed: finalData,
      method: finalMethod,
      crc32: crc,
      localOffset,
    });

    return this;
  }

  addText(name, content, compress = true) {
    return this.addFile(name, Buffer.from(content, 'utf8'), compress);
  }

  finalize() {
    const body = this.bodyWriter.toBuffer();
    const cdWriter = new ByteWriter();
    const cdOffset = body.length;

    for (const entry of this.entries) {
      const nameBuf = Buffer.from(entry.name, 'utf8');
      cdWriter
        .writeU32le(SIG_CENTRAL)
        .writeU16le(20)
        .writeU16le(20)
        .writeU16le(FLAG_UTF8_NAME)
        .writeU16le(entry.method)
        .writeU16le(0)
        .writeU16le(0)
        .writeU32le(entry.crc32)
        .writeU32le(entry.compressed.length)
        .writeU32le(entry.data.length)
        .writeU16le(nameBuf.length)
        .writeU16le(0)
        .writeU16le(0)
        .writeU16le(0)
        .writeU16le(0)
        .writeU32le(0)
        .writeU32le(entry.localOffset)
        .writeBytes(nameBuf);
    }

    const cd = cdWriter.toBuffer();

    const eocdWriter = new ByteWriter();
    eocdWriter
      .writeU32le(SIG_EOCD)
      .writeU16le(0)
      .writeU16le(0)
      .writeU16le(this.entries.length)
      .writeU16le(this.entries.length)
      .writeU32le(cd.length)
      .writeU32le(cdOffset)
      .writeU16le(0);

    return Buffer.concat([body, cd, eocdWriter.toBuffer()]);
  }
}
