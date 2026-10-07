/**
 * Bundled, OFL-licensed fonts this project ships locally so PDF export never
 * needs the network or a caller-supplied file for them. Two different
 * reasons a name ends up here:
 *
 *  - Arimo (Arial), Tinos (Times New Roman), Cousine (Courier New), Gelasio
 *    (Georgia) are metric-compatible SUBSTITUTES — the requested name is a
 *    proprietary font this project can never legally redistribute, so these
 *    stand in for it under the proprietary name.
 *  - Poppins, Roboto, Open Sans, Montserrat, Inter, Rubik, and Source Sans 3
 *    are the REAL fonts themselves, bundled under their own names — all
 *    freely OFL-licensed already, no substitute needed; they're here purely
 *    because they're among the most commonly-requested web fonts, worth
 *    resolving offline by default rather than only via a document's own
 *    declared stylesheets.
 *
 * Every one of the 6 latter fonts (Poppins is the only exception) is
 * distributed by Google Fonts as a VARIABLE font (one file, a continuous
 * weight axis) — not as separate static Regular/Bold/Italic/BoldItalic
 * files. The 4 `.ttf` files bundled here per family were produced OFFLINE,
 * once, from those variable sources via fontTools' `varLib.instancer`
 * (pinning `wght` to 400/700 and any other axis — Roboto/Open Sans's `wdth`,
 * Inter's `opsz` — to its default), NOT by any code in this repository or at
 * runtime: this project has no variable-font instantiation of its own, and
 * deliberately doesn't attempt to add one just for this — see this session's
 * own conversation history for why that was scoped out as disproportionate
 * effort for six specific fonts. The result is four ordinary static font
 * files, indistinguishable at load time from a font that was always static.
 *
 * License text for each is alongside the .ttf files in ./builtin/*-OFL.txt.
 *
 * Any other requested font name with no real registered glyph data (a
 * web-safe name like Verdana with no freely-redistributable clone, or any
 * other unresolved custom name) has no approximation bucket to fall into —
 * PDF text with no real embedded font just draws as plain Helvetica (see
 * exporter/pdf/fonts.js's pickBaseFont). Helvetica itself needs no bundled
 * entry either: it's already a true zero-embedding PDF Standard-14 font, so
 * embedding a substitute for it would only add file size for no fidelity
 * gain.
 */
import { readFile } from 'node:fs/promises';
import { decodeFontFile } from './resolver.js';
import { collectFontFamiliesUsed } from './collect-font-usages.js';

const BUILTIN_DIR = new URL('./builtin/', import.meta.url);

const BUILTIN_FONT_FAMILY = {
  arial: 'Arimo',
  'times new roman': 'Tinos',
  'courier new': 'Cousine',
  georgia: 'Gelasio',
  poppins: 'Poppins',
  roboto: 'Roboto',
  'open sans': 'OpenSans',
  montserrat: 'Montserrat',
  inter: 'Inter',
  rubik: 'Rubik',
  'source sans 3': 'SourceSans3',
  'source sans pro': 'SourceSans3', // Source Sans 3's own predecessor name — still commonly used in existing HTML
};

const VARIANTS = [
  { weight: 400, italic: false, suffix: 'Regular' },
  { weight: 700, italic: false, suffix: 'Bold' },
  { weight: 400, italic: true, suffix: 'Italic' },
  { weight: 700, italic: true, suffix: 'BoldItalic' },
];

/**
 * Registers whichever bundled fonts the document actually references
 * (collectFontFamiliesUsed — same scoping reasoning as
 * stylesheet-resolver.js's own network discovery: a document that never
 * asks for Georgia shouldn't pay to read Gelasio's four files off disk).
 * Local files only, so this runs unconditionally wherever `fonts/resolver.js`
 * is being fed at all.
 *
 * Never overrides an already-registered exact weight/italic combo: a page's
 * own discovered `@font-face` (which runs before this) always wins over a
 * bundled font here — same "explicit/real beats generic assumption"
 * precedent `resolveFontFamilyStack` itself now applies at lookup time.
 */
export async function loadBuiltinSubstituteFonts(document, registry, warn) {
  const familiesUsed = collectFontFamiliesUsed(document);
  if (familiesUsed.size === 0) return registry;

  const jobs = Object.entries(BUILTIN_FONT_FAMILY)
    .filter(([cssName]) => familiesUsed.has(cssName))
    .flatMap(([cssName, bundledFamily]) => VARIANTS.map((variant) => ({ cssName, bundledFamily, variant })));

  await Promise.all(jobs.map(async ({ cssName, bundledFamily, variant }) => {
    const { weight, italic, suffix } = variant;
    if (registry.hasExact(cssName, weight, italic)) return;
    try {
      const raw = await readFile(new URL(`${bundledFamily}-${suffix}.ttf`, BUILTIN_DIR));
      const { sfnt, buffer } = decodeFontFile(raw);
      registry.register(cssName, weight, italic, sfnt, buffer);
    } catch (err) {
      warn?.('BUILTIN_FONT_LOAD_FAILED', `Cannot load bundled font "${cssName}" (weight ${weight}${italic ? ' italic' : ''}): ${err instanceof Error ? err.message : String(err)}`);
    }
  }));

  return registry;
}
