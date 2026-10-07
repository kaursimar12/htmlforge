/**
 * Regenerates exporter/shared/font-metrics.js's per-character advance-width
 * tables from the real bundled substitute fonts (src/fonts/builtin/) instead
 * of hand-typed/scaled numbers — run after changing which font stands in for
 * which CSS name, or if the bundled .ttf files themselves are ever updated.
 * Not part of the runtime build: font-metrics.js stays pure sync/zero-I/O,
 * consuming only the plain object literals this prints, never reading a
 * font file itself.
 *
 * Usage: node scripts/generate-font-metrics.mjs
 */
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { parseSfnt } from '../src/fonts/sfnt-parser.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const BUILTIN_DIR = path.join(__dirname, '../src/fonts/builtin');

// Same row grouping as the tables this replaces, purely so the generated
// output is easy to eyeball/diff against the previous hand-typed version.
const ROWS = [
  [' ', '!', '"', '#', '$', '%', '&', "'", '(', ')', '*', '+', ',', '-', '.', '/'],
  ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'],
  [':', ';', '<', '=', '>', '?', '@'],
  [
    'A',
    'B',
    'C',
    'D',
    'E',
    'F',
    'G',
    'H',
    'I',
    'J',
    'K',
    'L',
    'M',
    'N',
    'O',
    'P',
    'Q',
    'R',
    'S',
    'T',
    'U',
    'V',
    'W',
    'X',
    'Y',
    'Z',
  ],
  ['[', '\\', ']', '^', '_', '`'],
  [
    'a',
    'b',
    'c',
    'd',
    'e',
    'f',
    'g',
    'h',
    'i',
    'j',
    'k',
    'l',
    'm',
    'n',
    'o',
    'p',
    'q',
    'r',
    's',
    't',
    'u',
    'v',
    'w',
    'x',
    'y',
    'z',
  ],
  ['{', '|', '}', '~'],
];

const IDENT_SAFE = /^[A-Za-z_$][A-Za-z0-9_$]*$/;

function keyFor(ch) {
  return IDENT_SAFE.test(ch) ? ch : JSON.stringify(ch);
}

async function widthsForFont(filename) {
  const buffer = await readFile(path.join(BUILTIN_DIR, filename));
  const sfnt = parseSfnt(buffer);
  const scale = 1000 / sfnt.unitsPerEm;
  const table = {};
  for (const row of ROWS) {
    for (const ch of row) {
      const gid = sfnt.glyphIdForCodepoint(ch.codePointAt(0));
      table[ch] =
        gid === undefined ? undefined : Math.round(sfnt.advanceWidthForGlyph(gid) * scale);
    }
  }
  return table;
}

function formatTable(constName, table) {
  const lines = ROWS.map(
    (row) => '  ' + row.map((ch) => `${keyFor(ch)}: ${table[ch]}`).join(', ') + ',',
  );
  return `const ${constName} = {\n${lines.join('\n')}\n};`;
}

const FONTS = [
  { constName: 'HELVETICA_WIDTHS', file: 'Arimo-Regular.ttf' },
  { constName: 'HELVETICA_BOLD_WIDTHS', file: 'Arimo-Bold.ttf' },
  { constName: 'TIMES_WIDTHS', file: 'Tinos-Regular.ttf' },
  { constName: 'TIMES_BOLD_WIDTHS', file: 'Tinos-Bold.ttf' },
  { constName: 'COURIER_WIDTHS', file: 'Cousine-Regular.ttf' },
  { constName: 'COURIER_BOLD_WIDTHS', file: 'Cousine-Bold.ttf' },
  { constName: 'GEORGIA_WIDTHS', file: 'Gelasio-Regular.ttf' },
  { constName: 'GEORGIA_BOLD_WIDTHS', file: 'Gelasio-Bold.ttf' },
  { constName: 'CALIBRI_WIDTHS', file: 'Carlito-Regular.ttf' },
  { constName: 'CALIBRI_BOLD_WIDTHS', file: 'Carlito-Bold.ttf' },
];

for (const { constName, file } of FONTS) {
  const table = await widthsForFont(file);
  const missing = Object.entries(table)
    .filter(([, v]) => v === undefined)
    .map(([k]) => k);
  if (missing.length) console.error(`// WARNING: ${file} has no glyph for: ${missing.join(' ')}`);
  console.log(
    `// Generated from ${file} via scripts/generate-font-metrics.mjs — do not hand-edit.`,
  );
  console.log(formatTable(constName, table));
  console.log();
}
