/**
 * Pure PDF content-stream operator builder — the PDF analogue of
 * `utils/xml-builder.js`'s `XmlBuilder`, but emitting operator syntax
 * instead of tags. Has no opinion on layout/page geometry; `page-flow.js`
 * and `blocks.js` decide *what* to draw and *where*, this only knows how to
 * spell it correctly as PDF content-stream text.
 */

// PDF numbers in content streams are plain decimal ASCII with no exponent
// notation — rounding to 3 decimal places and trimming trailing zeros keeps
// output compact without losing meaningful sub-point precision.
function num(n) {
  const rounded = Math.round(n * 1000) / 1000;
  return String(rounded);
}

// Escapes the three characters PDF's literal-string syntax `(...)` requires
// escaping — backslash and unbalanced parens would otherwise corrupt the
// stream's own structure.
export function pdfString(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
}

// Every standard font uses /Encoding /WinAnsiEncoding (Windows-1252). Bytes
// 0x00-0xff of WinAnsi match raw Unicode code points *except* the 0x80-0x9f
// range, which Windows-1252 repurposes for common punctuation (smart
// quotes, en/em dash, bullet, ellipsis, trademark...) that real HTML
// commonly contains but whose real Unicode code points (U+2018-U+2122) sit
// far outside 0-255. Without this table, the final content-stream buffer
// build (`Buffer.from(text, 'latin1')`) would silently truncate each such
// character to its low byte instead of the correct WinAnsi byte — e.g. a
// bullet U+2022 truncates to 0x22, a plain double-quote, not a bullet.
const WINANSI_HIGH_CODEPOINTS = {
  0x20ac: 0x80, 0x201a: 0x82, 0x0192: 0x83, 0x201e: 0x84, 0x2026: 0x85,
  0x2020: 0x86, 0x2021: 0x87, 0x02c6: 0x88, 0x2030: 0x89, 0x0160: 0x8a,
  0x2039: 0x8b, 0x0152: 0x8c, 0x017d: 0x8e, 0x2018: 0x91, 0x2019: 0x92,
  0x201c: 0x93, 0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97,
  0x02dc: 0x98, 0x2122: 0x99, 0x0161: 0x9a, 0x203a: 0x9b, 0x0153: 0x9c,
  0x017e: 0x9e, 0x0178: 0x9f,
};

/**
 * Re-encodes `text` so every character is already a single byte (0-255)
 * matching its real WinAnsiEncoding value, ready for the later
 * `Buffer.from(fullContentStream, 'latin1')` conversion to round-trip
 * correctly. A code point with no WinAnsi representation at all (CJK,
 * Arabic, Cyrillic, emoji — standard-14 fonts have no embedded glyphs for
 * these regardless of encoding) becomes `?` and sets `hasUnsupported` —
 * `page-flow.js` surfaces this as a once-per-document warning, the real,
 * honest content-fidelity ceiling of using non-embedded standard fonts.
 */
export function encodeWinAnsi(text) {
  let out = '';
  let hasUnsupported = false;
  for (const ch of text) {
    const codePoint = ch.codePointAt(0);
    if (codePoint <= 0xff) { out += ch; continue; }
    const mapped = WINANSI_HIGH_CODEPOINTS[codePoint];
    if (mapped != null) { out += String.fromCharCode(mapped); continue; }
    out += '?';
    hasUnsupported = true;
  }
  return { text: out, hasUnsupported };
}

export class ContentStreamBuilder {
  #ops = [];

  save() { this.#ops.push('q'); return this; }
  restore() { this.#ops.push('Q'); return this; }

  translate(x, y) { this.#ops.push(`1 0 0 1 ${num(x)} ${num(y)} cm`); return this; }

  /** Full 2x2-plus-translation matrix concat — an image XObject's own unit
   * square (0,0)-(1,1) needs both a scale (to its real drawn size) and a
   * translate (to its drawn position) in one `cm` before `Do` paints it. */
  concatMatrix(a, b, c, d, e, f) { this.#ops.push(`${num(a)} ${num(b)} ${num(c)} ${num(d)} ${num(e)} ${num(f)} cm`); return this; }

  setFillColorRGB(r, g, b) { this.#ops.push(`${num(r)} ${num(g)} ${num(b)} rg`); return this; }
  setStrokeColorRGB(r, g, b) { this.#ops.push(`${num(r)} ${num(g)} ${num(b)} RG`); return this; }
  setLineWidth(w) { this.#ops.push(`${num(w)} w`); return this; }

  rect(x, y, w, h) { this.#ops.push(`${num(x)} ${num(y)} ${num(w)} ${num(h)} re`); return this; }
  fill() { this.#ops.push('f'); return this; }
  stroke() { this.#ops.push('S'); return this; }
  fillAndStroke() { this.#ops.push('B'); return this; }

  moveTo(x, y) { this.#ops.push(`${num(x)} ${num(y)} m`); return this; }
  lineTo(x, y) { this.#ops.push(`${num(x)} ${num(y)} l`); return this; }
  curveTo(x1, y1, x2, y2, x3, y3) { this.#ops.push(`${num(x1)} ${num(y1)} ${num(x2)} ${num(y2)} ${num(x3)} ${num(y3)} c`); return this; }
  closePath() { this.#ops.push('h'); return this; }

  /** A rectangle with `r`-radius rounded corners, built from 4 straight edges
   * + 4 quarter-circle Bézier arcs — PDF's `re` operator has no native
   * rounded-rect primitive, unlike DrawingML's `roundRect` preset geometry.
   * `k` is the standard cubic-Bézier constant for approximating a quarter
   * circle (4/3 * (sqrt(2)-1)), the same technique most vector toolkits use. */
  roundedRect(x, y, w, h, r) {
    const k = 0.5523 * r;
    this.moveTo(x + r, y)
      .lineTo(x + w - r, y)
      .curveTo(x + w - r + k, y, x + w, y + r - k, x + w, y + r)
      .lineTo(x + w, y + h - r)
      .curveTo(x + w, y + h - r + k, x + w - r + k, y + h, x + w - r, y + h)
      .lineTo(x + r, y + h)
      .curveTo(x + r - k, y + h, x, y + h - r + k, x, y + h - r)
      .lineTo(x, y + r)
      .curveTo(x, y + r - k, x + r - k, y, x + r, y)
      .closePath();
    return this;
  }

  /** A full ellipse inscribed in the (x,y,w,h) box, via the same 4-arc
   * Bézier technique as roundedRect — used for `border-radius:50%+`
   * ("circular"), matching DrawingML's `ellipse` preset geometry. */
  ellipse(x, y, w, h) {
    const rx = w / 2;
    const ry = h / 2;
    const cx = x + rx;
    const cy = y + ry;
    const kx = 0.5523 * rx;
    const ky = 0.5523 * ry;
    this.moveTo(cx + rx, cy)
      .curveTo(cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry)
      .curveTo(cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy)
      .curveTo(cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry)
      .curveTo(cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy)
      .closePath();
    return this;
  }

  beginText() { this.#ops.push('BT'); return this; }
  endText() { this.#ops.push('ET'); return this; }
  setFont(resourceName, sizePt) { this.#ops.push(`/${resourceName} ${num(sizePt)} Tf`); return this; }
  setTextPosition(x, y) { this.#ops.push(`1 0 0 1 ${num(x)} ${num(y)} Tm`); return this; }
  showText(text) { this.#ops.push(`(${pdfString(text)}) Tj`); return this; }

  /** Shows text for a Type0/Identity-H embedded (CID) font — `hexDigits` is
   * an already-built ASCII hex string, 4 hex digits (2 bytes, big-endian) per
   * glyph ID, e.g. "004100420043" for 3 glyphs. Unlike `showText`'s literal
   * `(...)` syntax (one BYTE per character, WinAnsi-encoded), a CID font's
   * codes are 2 BYTES per glyph — PDF's hex-string syntax `<...>` is how the
   * content stream spells that, and because hex digits are themselves plain
   * ASCII, this string survives the later whole-stream
   * `Buffer.from(ops, 'latin1')` conversion with no special-casing needed. */
  showHexText(hexDigits) { this.#ops.push(`<${hexDigits}> Tj`); return this; }

  drawImage(resourceName) { this.#ops.push(`/${resourceName} Do`); return this; }

  toString() { return this.#ops.join('\n') + '\n'; }
}
