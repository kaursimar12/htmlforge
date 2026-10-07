/**
 * Walks the whole IDM tree collecting every distinct font-family NAME
 * referenced anywhere in the document (each entry of every styledSpan's own
 * comma-separated `font-family` stack, plus the document-wide
 * `body{font-family}` default) — used to scope network-fetched stylesheet
 * discovery (fonts/stylesheet-resolver.js) to families the document
 * actually asks for, so converting a page that merely LINKS a Google Fonts
 * stylesheet declaring 20 weights doesn't download all 20 when the page
 * text only ever uses 2 of them.
 *
 * Returns a Map, lowercased key -> first-seen ORIGINAL casing, not a Set —
 * every existing consumer only ever calls `.has()`/`.size` (a Map behaves
 * identically to a Set for those), but google-fonts-fallback.js specifically
 * needs the original casing back: Google's own CSS2 API is case-sensitive
 * (`family=Lato` resolves, `family=lato` 400s), so the lowercased form this
 * module normalizes to everywhere else isn't usable for that one caller.
 *
 * Same generic recursive-walk shape as assets/resolver.js's own
 * `collectImages`/`walk` (a `.children`/`.items` fallback rather than an
 * enumerated per-node-type switch, plus the same explicit table
 * head+rows+cells handling) — deliberately not shared code with it since
 * this collects family names rather than mutating image nodes in place, but
 * the shape of "what counts as a child to recurse into" needs to stay
 * identical or a node type images.js already knows to recurse into (e.g. a
 * table cell's block content) would silently go unwalked here.
 */
function stripQuotes(name) {
  return name.trim().replace(/^['"]|['"]$/g, '');
}

function addStack(fontFamilyCss, out) {
  if (!fontFamilyCss) return;
  for (const entry of fontFamilyCss.split(',')) {
    const original = stripQuotes(entry);
    if (!original) continue;
    const lower = original.toLowerCase();
    if (!out.has(lower)) out.set(lower, original); // first-seen casing wins
  }
}

function walk(nodes, out) {
  for (const node of nodes ?? []) {
    if (!node) continue;
    if (node.type === 'styledSpan' && node.style?.fontFamily) addStack(node.style.fontFamily, out);
    if (node.type === 'table') {
      for (const row of [...(node.head ?? []), ...node.rows]) {
        for (const cell of row.cells) walk(cell.children ?? [], out);
      }
      continue;
    }
    if (node.children) walk(node.children, out);
    if (node.items) for (const item of node.items) walk(item.children, out);
  }
}

export function collectFontFamiliesUsed(document) {
  const out = new Map();
  addStack(document.metadata?.fontFamily, out);
  walk(document.children, out);
  return out;
}

// Capped so a pathologically large document can't make this an unbounded
// scan — a representative sample of the document's own script mix (Latin,
// Cyrillic, CJK, ...) is what this exists for, not exhaustive coverage; a
// few hundred distinct codepoints is already far more than enough to tell
// stylesheet-resolver.js which @font-face Unicode-range subset the document
// actually needs (see its own doc comment on why "always assume Latin" was
// wrong for non-Latin-primary documents).
const MAX_CODEPOINTS = 500;

function walkText(nodes, out) {
  for (const node of nodes ?? []) {
    if (!node || out.size >= MAX_CODEPOINTS) return;
    if (node.type === 'text' && node.value) {
      for (const ch of node.value) {
        out.add(ch.codePointAt(0));
        if (out.size >= MAX_CODEPOINTS) return;
      }
      continue;
    }
    if (node.type === 'table') {
      for (const row of [...(node.head ?? []), ...node.rows]) {
        for (const cell of row.cells) walkText(cell.children ?? [], out);
      }
      continue;
    }
    if (node.children) walkText(node.children, out);
    if (node.items) for (const item of node.items) walkText(item.children, out);
  }
}

/**
 * Every distinct Unicode codepoint (capped, see MAX_CODEPOINTS) appearing
 * anywhere in the document's actual text content — used by
 * stylesheet-resolver.js to pick whichever `@font-face` Unicode-range subset
 * rule the document's own script mix actually needs, instead of assuming
 * Latin-primary text for every document.
 */
export function collectCodepointsUsed(document) {
  const out = new Set();
  walkText(document.children, out);
  return out;
}
