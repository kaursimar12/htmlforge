/**
 * Exporter contract. An exporter takes an IDM DocumentNode and produces the
 * target file as a raw Buffer. The core pipeline handles HTML parsing and IDM
 * normalisation, so exporters never touch raw HTML directly.
 */

export function createContext(options, warnings, assets, fonts) {
  return {
    options,
    warnings,
    assets,
    // Resolved once by convert() (src/index.js), same relationship `assets`
    // already has to the image resolver — only exporters with
    // `embedsAssets: true` (PDF, today) ever populate/read this; DOCX/PPTX
    // only ever need a font-family NAME decision, never real font bytes
    // (see fonts/resolver.js's own doc comment).
    fonts,
    warn(code, message, severity = 'warning') {
      warnings.push({ code, message, severity });
    },
  };
}
