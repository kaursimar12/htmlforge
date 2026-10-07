# html-parsor

> Universal HTML-to-document exporter. Clean HTML in, any format out.

`html-parsor` converts HTML into any document format you need — Word, Excel, PowerPoint, PDF, Markdown, CSV, plain text, or JSON. HTML is the only input. One API. One input format. Any output.

> **Status:** early scaffold (v0.1). All 8 exporters are implemented: **Markdown, CSV, TXT, JSON, DOCX, XLSX, PPTX, PDF**.

```
        Clean HTML
              │
              ▼
  Universal HTML Exporter
              │
              ▼
    Any Document Format
```

## Install

```bash
npm install html-parsor
```

## Usage

```ts
import { convert } from 'html-parsor';
import { writeFile } from 'node:fs/promises';

const html = `
  <h1>Quarterly Report</h1>
  <p>Summary paragraph with <strong>bold</strong> text.</p>
  <table>
    <tr><th>Item</th><th>Value</th></tr>
    <tr><td>Revenue</td><td>$1M</td></tr>
  </table>
`;

// Export to Word
const docx = await convert(html, { format: 'docx' });
await writeFile('report.docx', docx.buffer);

// Export to PDF
const pdf = await convert(html, { format: 'pdf' });
await writeFile('report.pdf', pdf.buffer);

// Export to Markdown (text output)
const md = await convert(html, { format: 'markdown' });
console.log(md.buffer.toString('utf8'));
```

### From a file path

```ts
const result = await convert('./report.html', { format: 'xlsx' });
```

### From a Buffer

```ts
const result = await convert(htmlBuffer, { format: 'csv' });
```

## Output contract

Every call returns the same shape:

```ts
interface ExportResult {
  buffer: Buffer;      // the exported file as raw bytes
  mimeType: string;    // e.g. "application/pdf"
  extension: string;   // e.g. ".pdf"
  metadata: Metadata;  // title, author, format, …
  warnings: Warning[]; // recoverable issues (skipped elements, missing images, …)
}
```

## Supported formats

| Format | `format` value | Extension | Notes |
|--------|---------------|-----------|-------|
| Microsoft Word | `"docx"` | `.docx` | Headings, paragraphs, lists, tables, images |
| Microsoft Excel | `"xlsx"` | `.xlsx` | Each `<table>` → one sheet |
| Microsoft PowerPoint | `"pptx"` | `.pptx` | Each `<section>` → one slide |
| PDF | `"pdf"` | `.pdf` | Full layout with page flow |
| Markdown | `"markdown"` | `.md` | GFM-compatible |
| CSV | `"csv"` | `.csv` | First `<table>` → RFC 4180 CSV |
| Plain text | `"txt"` | `.txt` | ASCII tables, ALL CAPS headings |
| JSON | `"json"` | `.json` | Document structure as typed JSON |

## Options

```ts
await convert(html, {
  format: 'docx',

  // Document metadata
  title: 'Q3 Report',
  author: 'Finance Team',
  language: 'en',

  // Layout (PDF, DOCX)
  pageSize: 'A4',           // 'A4' | 'Letter' | 'Legal'
  orientation: 'portrait',  // 'portrait' | 'landscape'
  margins: { top: 20, right: 20, bottom: 20, left: 20 }, // mm

  // Images
  embedImages: true,
  baseUrl: 'https://myapp.com', // resolve relative image URLs

  // XLSX-specific
  sheetName: 'Sheet1',

  // PPTX-specific
  slideSize: 'widescreen',  // 'widescreen' (16:9) | 'standard' (4:3)

  // Markdown-specific
  flavor: 'gfm',            // 'gfm' | 'commonmark'
});
```

## Architecture

```
HTML Input (string | Buffer | file path)
   │
   ▼
HTML Parser  →  Intermediate Document Model (IDM)
                          │
                          ▼
              Exporter Selection (by format)
                          │
        ┌────────┬────────┼────────┬────────┐
        ▼        ▼        ▼        ▼        ▼
      DOCX     XLSX     PPTX     PDF    Markdown / CSV / TXT / JSON
        │        │        │        │        │
        └────────┴────────┴────────┴────────┘
                          │
                          ▼
              ExportResult { buffer, mimeType, extension, metadata, warnings }
```

| Folder | Responsibility |
|--------|---------------|
| `src/html-parser/` | HTML string → Intermediate Document Model |
| `src/idm/` | IDM node type definitions |
| `src/normalizer/` | Structural clean-up before export |
| `src/exporter/markdown/` | IDM → Markdown |
| `src/exporter/csv/` | IDM → CSV |
| `src/exporter/txt/` | IDM → plain text |
| `src/exporter/json/` | IDM → JSON |
| `src/exporter/docx/` | IDM → DOCX (WordprocessingML) |
| `src/exporter/xlsx/` | IDM → XLSX (SpreadsheetML) |
| `src/exporter/pptx/` | IDM → PPTX (PresentationML) |
| `src/exporter/pdf/` | IDM → PDF 1.4 |
| `src/utils/` | ZIP writer, XML builder, binary utilities |

## Extending

Register a custom exporter:

```ts
import { registerExporter } from 'html-parsor';

registerExporter({
  id: 'rtf',
  extensions: ['.rtf'],
  mimeType: 'application/rtf',
  export(document, ctx) {
    // build and return a Buffer
  },
});
```

## Design principles

- **HTML is the only input** — no PDF, DOCX, images, or other formats as input
- **Zero runtime npm dependencies** — built entirely on Node.js core (zlib, fs, buffer, crypto)
- **One API, any output** — same `convert()` call regardless of target format
- **Full TypeScript** — complete types for all inputs, outputs, and options

## Development

```bash
npm install
npm run build       # bundle ESM + CJS + types (tsup)
npm test            # run unit tests (vitest)
npm run typecheck   # tsc --noEmit
npm run lint        # eslint
```

## License

MIT
