/**
 * Dedup key for a run/font's stylable properties (bold/italic/underline/
 * strikethrough/color/fontSizePt) — shared by styles.js's font registry
 * (`registerCellStyle` needs "have we already registered this exact font
 * combination") and rich-text.js's adjacent-run merging (`mergeAdjacent`
 * needs "do these two runs describe the same visual font"), since both are
 * the same six-field equality check.
 */
export function fontStyleKey({ bold, italic, underline, strikethrough, color, fontSizePt } = {}) {
  return JSON.stringify({ bold: !!bold, italic: !!italic, underline: !!underline, strikethrough: !!strikethrough, color, fontSizePt });
}
