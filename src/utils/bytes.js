/**
 * Minimal binary reader/writer over a {@link Buffer}.
 *
 * Used by the ZIP reader/writer to walk and produce byte structures without
 * any parsing dependency.
 */

export class ByteReader {
  #pos = 0;

  constructor(buffer) {
    this.buffer = buffer;
  }

  get offset() { return this.#pos; }
  get remaining() { return this.buffer.length - this.#pos; }
  get eof() { return this.#pos >= this.buffer.length; }

  seek(position) {
    if (position < 0 || position > this.buffer.length)
      throw new RangeError(`seek out of range: ${position} (len ${this.buffer.length})`);
    this.#pos = position;
    return this;
  }

  skip(n) { return this.seek(this.#pos + n); }

  u8() { this.#ensure(1); return this.buffer[this.#pos++]; }

  u16le() {
    this.#ensure(2);
    const v = this.buffer.readUInt16LE(this.#pos);
    this.#pos += 2;
    return v;
  }

  u32le() {
    this.#ensure(4);
    const v = this.buffer.readUInt32LE(this.#pos);
    this.#pos += 4;
    return v;
  }

  u16be() {
    this.#ensure(2);
    const v = this.buffer.readUInt16BE(this.#pos);
    this.#pos += 2;
    return v;
  }

  u32be() {
    this.#ensure(4);
    const v = this.buffer.readUInt32BE(this.#pos);
    this.#pos += 4;
    return v;
  }

  i16be() {
    this.#ensure(2);
    const v = this.buffer.readInt16BE(this.#pos);
    this.#pos += 2;
    return v;
  }

  bytes(n) {
    this.#ensure(n);
    const slice = this.buffer.subarray(this.#pos, this.#pos + n);
    this.#pos += n;
    return slice;
  }

  string(n, encoding = 'utf8') {
    return this.bytes(n).toString(encoding);
  }

  peekU32le() {
    this.#ensure(4);
    return this.buffer.readUInt32LE(this.#pos);
  }

  #ensure(n) {
    if (this.#pos + n > this.buffer.length)
      throw new RangeError(
        `read past end of buffer: need ${n} at ${this.#pos}, len ${this.buffer.length}`,
      );
  }
}

/** Growable byte writer — builds a Buffer incrementally. */
export class ByteWriter {
  #chunks = [];
  #length = 0;

  get length() { return this.#length; }

  writeU8(v) {
    const b = Buffer.allocUnsafe(1);
    b.writeUInt8(v, 0);
    return this.#push(b);
  }

  writeU16le(v) {
    const b = Buffer.allocUnsafe(2);
    b.writeUInt16LE(v, 0);
    return this.#push(b);
  }

  writeU32le(v) {
    const b = Buffer.allocUnsafe(4);
    b.writeUInt32LE(v, 0);
    return this.#push(b);
  }

  writeU16be(v) {
    const b = Buffer.allocUnsafe(2);
    b.writeUInt16BE(v, 0);
    return this.#push(b);
  }

  writeU32be(v) {
    const b = Buffer.allocUnsafe(4);
    b.writeUInt32BE(v, 0);
    return this.#push(b);
  }

  writeBytes(buf) {
    return this.#push(Buffer.isBuffer(buf) ? buf : Buffer.from(buf));
  }

  writeString(s, encoding = 'latin1') {
    return this.#push(Buffer.from(s, encoding));
  }

  toBuffer() { return Buffer.concat(this.#chunks); }

  #push(b) {
    this.#chunks.push(b);
    this.#length += b.length;
    return this;
  }
}

export function startsWith(buffer, signature) {
  if (buffer.length < signature.length) return false;
  for (let i = 0; i < signature.length; i++) {
    if (buffer[i] !== signature[i]) return false;
  }
  return true;
}
