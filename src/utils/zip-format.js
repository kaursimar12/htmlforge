/**
 * ZIP file format constants shared by the reader (zip.js) and writer
 * (zip-writer.js) — the same signatures/method codes/flags on both sides of
 * the same binary format.
 */

export const SIG_LOCAL = 0x04034b50;
export const SIG_CENTRAL = 0x02014b50;
export const SIG_EOCD = 0x06054b50;

export const METHOD_STORED = 0;
export const METHOD_DEFLATE = 8;

// General-purpose bit flag 11 ("Language encoding flag (EFS)") — set when a
// filename is UTF-8, as opposed to the legacy default (CP437/latin1).
export const FLAG_UTF8_NAME = 0x0800;
