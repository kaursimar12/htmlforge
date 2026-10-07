# Product Requirements Document (PRD)

# Universal HTML-to-Document Exporter (Node.js / NPM Package)

**Version:** 2.0  
**Status:** Draft  
**Target Platform:** Node.js (NPM Package)

---

# 1. Overview

## Problem Statement

HTML is the universal language of the web. It is produced by:

- Rich text editors (Notion, Quill, TipTap, ProseMirror)
- CMS platforms (WordPress, Contentful, Sanity)
- Email builders
- AI / LLM output
- Web scrapers
- Documentation tools

But the moment you need to deliver that content as a **real document** — a PDF report, a Word file, an Excel spreadsheet, a PowerPoint deck — you face a different library for every format, each with its own API, quirks, and output quality.

---

## Goal

Build a universal document export library whose **only input format is HTML**.

Regardless of the target format, every exporter takes clean HTML in and produces the requested document out.

```text
        Clean HTML
              │
              ▼
  Universal HTML Exporter
              │
              ▼
    Any Document Format
```

HTML becomes the single source of truth. Write once, export anywhere.

---

# 2. Vision

Instead of integrating a different library for every output format:

```text
HTML
  │
  ▼
Universal Exporter
  │
  ├── PDF
  ├── DOCX (Microsoft Word)
  ├── XLSX (Microsoft Excel)
  ├── PPTX (Microsoft PowerPoint)
  ├── Markdown
  ├── CSV
  ├── TXT
  ├── JSON
  └── XML
```

One API. One input format. Any output.

---

# 3. Design Principles

## 3.1 HTML Is the Only Input

The package accepts only:

- HTML string
- HTML file path
- HTML Buffer / Uint8Array

It will never accept PDF, DOCX, images, or other formats as input. Those are out of scope.

---

## 3.2 Preserve Document Structure

Example HTML input:

```html
<h1>Heading</h1>

<p>Paragraph</p>

<ul>
  <li>Bullet</li>
</ul>

<table>...</table>

<figure>
    <img src="..." />
    <figcaption>Caption</figcaption>
</figure>
```

When exported to DOCX becomes:

```
Heading      → Word Heading 1 style
Paragraph    → Word Normal paragraph
Bullet       → Word List Bullet style
Table        → Word table with borders
Figure       → Inline image with caption
```

When exported to Markdown becomes:

```markdown
# Heading

Paragraph

- Bullet

| col | col |
|-----|-----|

![Caption](image.png)
```

---

## 3.3 Semantic HTML In, High-Fidelity Documents Out

The exporter reads semantic HTML elements and maps them to the closest native concept in the target format:

| HTML Element | DOCX | XLSX | PPTX | Markdown | CSV |
|---|---|---|---|---|---|
| `<h1>`–`<h6>` | Heading 1–6 style | Sheet title | Slide title | `#`–`######` | — |
| `<p>` | Normal paragraph | Cell text | Text box | paragraph | — |
| `<ul>` / `<ol>` | List Bullet / List Number | Cell list | Bullet list | `-` / `1.` | — |
| `<table>` | Word table | Sheet rows | Table shape | GFM table | rows |
| `<img>` | Inline image | Embedded image | Image shape | `![]()` | — |
| `<strong>` | Bold run | Bold cell | Bold run | `**text**` | — |
| `<em>` | Italic run | Italic cell | Italic run | `*text*` | — |
| `<a>` | Hyperlink | Hyperlink cell | Hyperlink | `[text](href)` | — |
| `<code>` | Courier font run | Monospace cell | Code text box | `` `code` `` | — |
| `<pre>` | Code block style | Merged cell | Code text box | ```` ``` ```` | — |
| `<blockquote>` | Quote style | Indented cell | Quote text box | `> text` | — |

---

## 3.4 Preserve Maximum Information

Whenever possible retain:

- Headings and their hierarchy
- Paragraphs and inline formatting (bold, italic, underline, code)
- Tables with headers, merged cells, captions
- Lists (ordered and unordered, nested)
- Images with alt text and captions
- Hyperlinks
- Page structure (sections, articles, headers, footers)
- Metadata (title, author, language)
- Footnotes and references

---

# 4. Supported Output Formats

## Phase 1 — Text Formats (no binary encoding)

### Markdown

Extension: `.md`

Capabilities:
- ATX headings (`#`–`######`)
- Paragraphs
- Bold, italic, inline code
- Ordered and unordered lists (nested)
- GFM tables
- Fenced code blocks
- Blockquotes
- Links and images
- Thematic breaks

---

### CSV

Extension: `.csv`

Capabilities:
- First `<table>` in the HTML → CSV rows
- Multiple tables → multiple CSV files (or sheets)
- Proper RFC 4180 quoting for fields containing commas/newlines

---

### Plain Text

Extension: `.txt`

Capabilities:
- Headings → ALL CAPS or underlined with `===`/`---`
- Paragraphs → blank-line separated
- Lists → `-` prefix
- Tables → ASCII art table
- Links → `text (url)`

---

### JSON

Extension: `.json`

Capabilities:
- HTML structure → JSON document object
- Tables → arrays of objects (column headers as keys)
- Lists → JSON arrays
- Headings + content → nested sections

---

## Phase 2 — Office Open XML (OOXML)

All three Office formats are ZIP archives of XML — built from scratch using the existing `utils/zip.ts` and `utils/xml.ts` primitives.

### Microsoft Word

Extension: `.docx`

Capabilities:
- Headings → `w:pStyle` Heading1–Heading6
- Paragraphs → `w:p` / `w:r` runs
- Bold / italic / underline → `w:rPr` run properties
- Lists → `w:numPr` with `numbering.xml`
- Tables → `w:tbl` / `w:tr` / `w:tc`
- Images → embedded in `word/media/`, referenced via `w:drawing`
- Hyperlinks → `w:hyperlink` with `document.xml.rels`
- Footnotes → `word/footnotes.xml`
- Metadata → `docProps/core.xml` (title, author, dates)

---

### Microsoft Excel

Extension: `.xlsx`

Capabilities:
- Each `<table>` → one worksheet
- Column headers (`<th>`) → header row with bold formatting
- Data rows (`<td>`) → `xl/worksheets/sheetN.xml`
- Multiple tables → multiple sheets in `xl/workbook.xml`
- Shared strings → `xl/sharedStrings.xml`
- Basic number/date formatting → `xl/styles.xml`

---

### Microsoft PowerPoint

Extension: `.pptx`

Capabilities:
- Each `<section>` → one slide
- `<h1>`/`<h2>` → title placeholder `<p:ph type="title"/>`
- Paragraph content → text body `<p:txBody>`
- Lists → bullet shapes
- Images → slide picture shapes
- Speaker notes → `ppt/notesSlides/notesSlideN.xml`
- Slide order → `ppt/presentation.xml`

---

## Phase 3 — PDF

Extension: `.pdf`

Capabilities:
- Text layout from HTML block flow
- Bold / italic font variants
- Tables with borders
- Embedded images
- Hyperlinks
- Page breaks on `<hr class="page-break" />`
- Page numbers in footer
- Document metadata (title, author) in PDF info dict

Implementation tiers (same risk model as original):
1. Basic text + paragraphs + headings
2. Tables + images
3. Full layout with page flow

---

## Phase 4 — Image Export

Extensions: `.png`, `.jpg`

Render the HTML to an image using a headless renderer. Decision needed: allow headless browser dependency (Puppeteer/Playwright) or implement a basic rasterizer from scratch.

---

# 5. Input Contract

The `convert` function accepts HTML in three forms:

```typescript
type HtmlInput =
    | string        // HTML string or file path
    | Buffer        // raw HTML bytes
    | Uint8Array;   // raw HTML bytes
```

Auto-detection:
- If the string starts with `<` → treated as an HTML string
- If the string is a valid file path → reads the file
- Otherwise → treated as an HTML string

---

# 6. Output Contract

Every exporter returns the same interface:

```typescript
interface ExportResult {
    buffer: Buffer;       // the exported file as raw bytes
    mimeType: string;     // e.g. "application/pdf"
    extension: string;    // e.g. ".pdf"
    metadata: Metadata;
    warnings: Warning[];
}
```

---

## Metadata

```typescript
interface Metadata {
    title?: string;
    author?: string;
    createdAt?: Date;
    language?: string;
    pageCount?: number;
    format?: string;      // e.g. "docx"
}
```

---

## Warnings

Examples:
- Unsupported HTML element (skipped)
- Image src not resolvable
- Table too wide for page (truncated)
- Nested table flattened (not supported in target format)
- Font not available in target format

---

# 7. Public API

## Convert to a specific format

```typescript
import { convert } from "html-exporter";

const html = `
  <h1>Quarterly Report</h1>
  <p>Summary paragraph.</p>
  <table>
    <tr><th>Item</th><th>Value</th></tr>
    <tr><td>Revenue</td><td>$1M</td></tr>
  </table>
`;

const result = await convert(html, { format: "docx" });

await writeFile("report.docx", result.buffer);
```

---

## Convert from an HTML file

```typescript
const result = await convert("./report.html", { format: "pdf" });
```

---

## Convert a Buffer

```typescript
const result = await convert(htmlBuffer, { format: "xlsx" });
```

---

## With options

```typescript
const result = await convert(html, {
    format: "docx",
    title: "Q3 Report",
    author: "Finance Team",
    language: "en",
    embedImages: true,
    pageSize: "A4",
    margins: { top: 20, right: 20, bottom: 20, left: 20 },
});
```

---

## Force a specific exporter

```typescript
const result = await convert(html, {
    exporter: "markdown"
});
```

---

# 8. Internal Architecture

```text
HTML Input (string | Buffer | file path)
   │
   ▼
HTML Parser
   │  Parses semantic HTML into IDM
   ▼
Intermediate Document Model (IDM)
   │
   ▼
Exporter Selection (by format)
   │
   ├── markdown  → Markdown Exporter → .md Buffer
   ├── csv       → CSV Exporter      → .csv Buffer
   ├── txt       → TXT Exporter      → .txt Buffer
   ├── json      → JSON Exporter     → .json Buffer
   ├── docx      → DOCX Exporter     → .docx Buffer (ZIP + XML)
   ├── xlsx      → XLSX Exporter     → .xlsx Buffer (ZIP + XML)
   ├── pptx      → PPTX Exporter     → .pptx Buffer (ZIP + XML)
   └── pdf       → PDF Exporter      → .pdf Buffer
   │
   ▼
ExportResult { buffer, mimeType, extension, metadata, warnings }
```

---

# 9. Intermediate Document Model (IDM)

The HTML parser converts HTML into the IDM. Every exporter reads from the IDM. This decoupling means:

- Adding a new output format = implementing one new exporter, no HTML parsing changes
- IDM is format-agnostic: it has no concept of Word styles or PDF operators

Supported IDM nodes:

- `DocumentNode` — root
- `SectionNode` — maps from `<section>`, `<article>`, `<main>`
- `HeadingNode` (level 1–6) — maps from `<h1>`–`<h6>`
- `ParagraphNode` — maps from `<p>`
- `ListNode` (ordered / unordered) — maps from `<ol>` / `<ul>`
- `ListItemNode` — maps from `<li>`
- `TableNode` (with head/body) — maps from `<table>`
- `TableRowNode` — maps from `<tr>`
- `TableCellNode` (colspan/rowspan) — maps from `<td>` / `<th>`
- `FigureNode` — maps from `<figure>`
- `ImageNode` — maps from `<img>`
- `CodeNode` — maps from `<pre><code>`
- `QuoteNode` — maps from `<blockquote>`
- `FootnoteNode` — maps from footnote markup
- `PageBreakNode` — maps from `<hr class="page-break">`
- `ThematicBreakNode` — maps from `<hr>`

Inline nodes:

- `TextNode`
- `StrongNode` — `<strong>` / `<b>`
- `EmphasisNode` — `<em>` / `<i>`
- `CodeSpanNode` — `<code>`
- `LinkNode` — `<a href>`
- `LineBreakNode` — `<br>`

---

# 10. HTML Parser

The HTML parser is the single entry point. It reads HTML and produces an IDM tree.

Strategy:
- Use `utils/html-tokenizer.ts` (to be built in Phase 2) — a forgiving tokenizer + tree builder
- Handle void elements, raw-text elements (`script`/`style`), and auto-closing rules
- Strip noise: `<script>`, `<style>`, `<nav>`, inline event handlers, tracking tags
- Resolve relative image URLs against a configurable base URL

Input HTML normalization before parsing:
- Collapse excessive whitespace
- Remove empty elements that carry no content
- Unwrap meaningless wrapper `<div>` / `<span>` chains

---

# 11. Exporter Interface

Every exporter implements the same contract:

```typescript
interface Exporter {
    id: string;               // e.g. "docx"
    extensions: string[];     // e.g. [".docx"]
    mimeType: string;         // e.g. "application/vnd.openxmlformats..."
    export(
        document: DocumentNode,
        ctx: ExporterContext
    ): Buffer | Promise<Buffer>;
}
```

The `ExporterContext` provides:
- Resolved options (page size, margins, embed settings, language)
- Asset resolver (fetch image bytes from src URLs or data: URIs)
- Warning emitter

---

# 12. Asset Handling

Images in input HTML may be:

- External URLs: `<img src="https://example.com/photo.jpg" />`
- Data URIs: `<img src="data:image/png;base64,..." />`
- Local file paths: `<img src="./images/chart.png" />`

The asset resolver fetches / reads image bytes and makes them available to exporters as `Buffer`. Exporters embed them in the output according to the target format's conventions (e.g. `word/media/image1.png` in DOCX, `DCTDecode` stream in PDF).

Configurable:
- `embedImages: true` (default) — embed bytes in the output file
- `embedImages: false` — keep external URLs as-is where the format supports it

---

# 13. Exporter Configuration

```typescript
await convert(html, {

    format: "docx",

    // Document metadata
    title: "Report Title",
    author: "Author Name",
    language: "en",

    // Layout (for paged formats: PDF, DOCX)
    pageSize: "A4",              // "A4" | "Letter" | "Legal"
    orientation: "portrait",    // "portrait" | "landscape"
    margins: {
        top: 20,   // mm
        right: 20,
        bottom: 20,
        left: 20,
    },

    // Images
    embedImages: true,
    baseUrl: "https://myapp.com", // resolve relative image URLs

    // XLSX-specific
    sheetName: "Sheet1",

    // PPTX-specific
    slideSize: "widescreen",     // "widescreen" (16:9) | "standard" (4:3)

    // Markdown-specific
    flavor: "gfm",               // "gfm" | "commonmark"

});
```

---

# 14. Error Handling

## Recoverable (emit warning, continue)

- Unknown HTML element → skip or treat as paragraph
- Image fetch failed → placeholder or skip
- Table too wide for page → truncate with warning
- Nested table in a format that doesn't support it → flatten

## Fatal (throw)

- Invalid / unparseable HTML string
- Unsupported format specified
- Output write failed
- Required asset missing and `strictAssets: true`

---

# 15. Performance Requirements

- HTML documents up to 50 MB
- Tables with 100,000+ rows (CSV / XLSX — streaming write)
- PPTX with 500+ slides
- PDF with 1,000+ pages

Additional:
- Streaming output where the format allows (XLSX row-by-row, PDF page-by-page)
- Memory-efficient image handling (stream rather than fully buffer large images)
- Parallel exporter runs: `convertAll(html, ['pdf', 'docx', 'xlsx'])` runs all three concurrently

---

# 16. Testing Strategy

## Unit Tests

- HTML parser: every semantic element maps to the correct IDM node
- Every exporter: every IDM node maps to the correct output construct
- Asset resolver: URL, data URI, local path

## Integration Tests (Golden Snapshots)

- Input HTML → expected .docx binary snapshot
- Input HTML → expected .md text snapshot
- Input HTML → expected .csv text snapshot
- Round-trip: HTML → DOCX → (open and verify structure programmatically)

## Performance Tests

- Large table (100k rows) → XLSX within memory budget
- Large HTML document → PDF within time budget

## Failure Tests

- Malformed HTML → graceful recovery, not crash
- Unreachable image URL → warning emitted, export continues
- Unknown HTML elements → skipped with warning

---

# 17. Package Structure

```
src/

    index.ts                  ← public API: convert()

    html-parser/              ← HTML string → IDM
        index.ts
        tokenizer.ts          ← forgiving HTML tokenizer (Phase 2)
        tree-builder.ts       ← DOM tree construction
        dom-to-idm.ts         ← DOM → IDM mapping

    idm/                      ← Intermediate Document Model nodes + builders
        index.ts

    exporter/
        registry.ts           ← exporter selection + registration
        types.ts              ← Exporter interface + ExporterContext
        markdown/
        csv/
        txt/
        json/
        docx/
        xlsx/
        pptx/
        pdf/

    assets/                   ← image fetching + resolution
        resolver.ts

    utils/
        bytes.ts              ← binary reader (Phase 0 ✅)
        zip.ts                ← ZIP reader + writer (Phase 0 ✅, writer: Phase 2)
        xml.ts                ← XML parser + serializer (Phase 0 ✅, serializer: Phase 2)
        magic.ts              ← format detection (Phase 0 ✅)
        html.ts               ← HTML escaping helpers

    types/
        index.ts              ← ExportResult, Metadata, Warning, ExportOptions

tests/

benchmarks/

examples/

docs/
```

---

# 18. Future Roadmap

- `convertAll(html, formats[])` — export to multiple formats in one call
- Template support: inject HTML content into a DOCX/PPTX template
- Style mapping: custom CSS → DOCX styles
- Table of contents auto-generation
- Header / footer templating (page numbers, document title)
- Mail merge: HTML template + JSON data → batch DOCX export
- Streaming API: pipe large HTML in, pipe output bytes out
- Plugin SDK: register custom exporters
- CLI: `html-export report.html --format pdf`
- Browser build (WASM) for client-side export

---

# 19. Success Metrics

| Metric | Target |
|---|---|
| Supported Output Formats | 8+ |
| Structure Fidelity | >95% of HTML elements correctly mapped |
| Maximum Input Size | 50 MB HTML |
| Image Embedding | All common formats (PNG, JPEG, GIF, WebP) |
| Streaming Output | Yes (XLSX, PDF) |
| Zero Runtime Dependencies | Yes (Node core only) |
| TypeScript Support | Full |
| Unit Test Coverage | >90% |

---

# 20. Long-Term Vision

This package becomes the **standard HTML export layer** for any application that produces HTML content and needs to deliver it as a real document.

```text
                 HTML
                  │
                  ▼
       Universal HTML Exporter
                  │
    ┌─────────────┼──────────────┐
    ▼             ▼              ▼
  DOCX          XLSX           PDF
(Word)        (Excel)        (Print)

    ▼             ▼              ▼
  PPTX        Markdown         CSV
(Slides)    (Developers)    (Data)
```

Write your content once as HTML. Deliver it anywhere.
