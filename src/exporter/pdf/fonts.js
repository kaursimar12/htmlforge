/**
 * Standard-14 font selection — PDF viewers guarantee these fourteen base
 * fonts exist with no embedding required, which is the entire reason a
 * from-scratch PDF writer can skip building/embedding real font files at
 * all for a run it has no real glyph data for.
 */
const HELVETICA_FAMILY = { plain: 'Helvetica', bold: 'Helvetica-Bold', italic: 'Helvetica-Oblique', boldItalic: 'Helvetica-BoldOblique' };
const COURIER_FAMILY = { plain: 'Courier', bold: 'Courier-Bold', italic: 'Courier-Oblique', boldItalic: 'Courier-BoldOblique' };

/**
 * Picks one of the 14 standard base font names for a run with no real
 * embedded glyph data (see page-flow.js's resolveTextFont/useEmbeddedFont for
 * the path that DOES have real glyphs). `code` (a CodeSpanNode/CodeNode, or a
 * styledSpan naming a mono family) always wins and forces Courier regardless
 * of bold/italic — matches the existing DOCX/PPTX "code is always monospace"
 * convention. Every other case is always Helvetica: this used to guess Times/
 * Courier from the declared family NAME via a regex, but that guess was
 * simply wrong whenever the name didn't happen to contain a recognized
 * substring (Garamond/Palatino/Bookman are serif fonts that don't contain
 * "serif"; Lucida Console is monospace but doesn't contain "mono"), so
 * guessing a family from the name was removed entirely rather than patched —
 * Helvetica is always the fallback for anything not actually embedded.
 */
export function pickBaseFont({ bold = false, italic = false, code = false } = {}) {
  const family = code ? COURIER_FAMILY : HELVETICA_FAMILY;
  if (bold && italic) return family.boldItalic;
  if (bold) return family.bold;
  if (italic) return family.italic;
  return family.plain;
}

/**
 * Assigns `/F1`, `/F2`... resource names to whichever base fonts actually
 * get used while walking the document — standard fonts cost nothing extra
 * to declare, but only declaring what's referenced keeps the `/Resources
 * /Font` dict (and the font object list `pdf-writer.js` allocates numbers
 * for) proportional to the document's own real font usage rather than
 * always emitting all 14.
 */
export class FontRegistry {
  #resourceNameByBaseFont = new Map();

  use(baseFont) {
    let name = this.#resourceNameByBaseFont.get(baseFont);
    if (!name) {
      name = `F${this.#resourceNameByBaseFont.size + 1}`;
      this.#resourceNameByBaseFont.set(baseFont, name);
    }
    return name;
  }

  /** `[{ baseFont, resourceName }]` in first-use order — stable iteration
   * order matters since pdf-writer.js allocates sequential object numbers
   * from this list. */
  entries() {
    return [...this.#resourceNameByBaseFont.entries()].map(([baseFont, resourceName]) => ({ baseFont, resourceName }));
  }
}
