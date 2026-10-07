import { describe, expect, it } from 'vitest';
import { ByteReader, startsWith } from '../src/utils/bytes.js';

describe('ByteReader', () => {
  it('reads integers of each width/endianness and advances the cursor', () => {
    const buf = Buffer.from([0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07]);
    const r = new ByteReader(buf);
    expect(r.u8()).toBe(0x01);
    expect(r.u16le()).toBe(0x0302);
    expect(r.offset).toBe(3);
    expect(r.u16be()).toBe(0x0405);
    expect(r.remaining).toBe(2);
  });

  it('reads 32-bit values both ways', () => {
    const buf = Buffer.from([0x78, 0x56, 0x34, 0x12]);
    expect(new ByteReader(buf).u32le()).toBe(0x12345678);
    expect(new ByteReader(buf).u32be()).toBe(0x78563412);
  });

  it('seek/skip and peek without consuming', () => {
    const buf = Buffer.from([0xaa, 0xbb, 0xcc, 0xdd]);
    const r = new ByteReader(buf);
    expect(r.peekU32le()).toBe(0xddccbbaa);
    expect(r.offset).toBe(0);
    r.skip(1);
    expect(r.u8()).toBe(0xbb);
    r.seek(0);
    expect(r.u8()).toBe(0xaa);
  });

  it('reads strings and raw byte views', () => {
    const buf = Buffer.from('PKhello', 'latin1');
    const r = new ByteReader(buf);
    expect(r.bytes(2).toString('latin1')).toBe('PK');
    expect(r.string(5)).toBe('hello');
  });

  it('throws when reading past the end', () => {
    const r = new ByteReader(Buffer.from([0x00]));
    expect(() => r.u32le()).toThrow(/past end/);
  });
});

describe('startsWith', () => {
  it('matches a leading signature', () => {
    expect(startsWith(Buffer.from([0x25, 0x50, 0x44, 0x46]), [0x25, 0x50])).toBe(true);
    expect(startsWith(Buffer.from([0x25, 0x50]), [0x25, 0x51])).toBe(false);
    expect(startsWith(Buffer.from([0x25]), [0x25, 0x50])).toBe(false);
  });
});
