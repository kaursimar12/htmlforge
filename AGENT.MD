# Reference Package Research — HTML → Any File Exporter

For each output format, this lists existing npm packages that convert **HTML → that format**.
We do **not** depend on these at runtime — we build from scratch.
These are here so we know what to study when implementing each exporter.

---

## HTML → PDF

| Package | Purpose |
| --- | --- |
| puppeteer | Headless Chrome HTML → PDF |
| playwright | Cross-browser HTML → PDF |
| wkhtmltopdf | WebKit-based HTML → PDF |
| html-pdf-node | Puppeteer-based HTML → PDF |
| pdfkit | Programmatic PDF generation |
| pdf-lib | PDF object model reference |
| jsPDF | Client-side PDF generation |

---

## HTML → DOCX

| Package | Purpose |
| --- | --- |
| html-to-docx | HTML → DOCX buffer |
| html-docx-js | HTML → DOCX via AltChunk |
| docx | Generate DOCX / WordprocessingML |
| officegen | OOXML boilerplate reference |
| docxtemplater | Template-based DOCX generation |

---

## HTML → XLSX

| Package | Purpose |
| --- | --- |
| node-html-to-xlsx | HTML tables → XLSX |
| xlsx (SheetJS) | HTML tables + XLSX writing |
| exceljs | Excel sheets, styles, workbooks |
| html-table-extractor | Extract HTML tables as JSON |

---

## HTML → PPTX

| Package | Purpose |
| --- | --- |
| pptxgenjs | Generate PPTX / PresentationML |
| officegen | PPTX boilerplate XML reference |

---

## HTML → Markdown

| Package | Purpose |
| --- | --- |
| turndown | HTML → Markdown |
| node-html-markdown | Fast HTML → Markdown |
| showdown | Markdown ↔ HTML |
| to-markdown | Older HTML → Markdown rules |
| rehype-remark | Convert HTML AST → Markdown AST (unified bridge) |
| remark-stringify | Serialize Markdown AST → Markdown string |

---

## HTML → CSV

| Package | Purpose |
| --- | --- |
| html-table-to-json | HTML tables → JSON arrays |
| json2csv | JSON → RFC 4180 CSV |
| fast-csv | Streaming CSV writer |
| papaparse | CSV writing/parsing |

---

## HTML → Plain Text

| Package | Purpose |
| --- | --- |
| html-to-text | HTML → readable plain text |
| striptags | Strip HTML tags |

---

## HTML → JSON

| Package | Purpose |
| --- | --- |
| htmlparser2 + domhandler | DOM tree to JSON |
| cheerio | Extract structured JSON |

---

## HTML → XML

| Package | Purpose |
| --- | --- |
| htmlparser2 | Parse HTML for XML output |
| rehype | HTML AST pipeline |
| rehype-stringify | AST → XML-like HTML |

---

## HTML → Image (PNG / JPEG)

| Package | Purpose |
| --- | --- |
| puppeteer | HTML screenshot → PNG/JPEG |
| playwright | Cross-browser HTML screenshots |
| wkhtmltoimage | WebKit HTML → image |
| node-html-to-image | Puppeteer wrapper HTML → image |

---

## HTML Parsing (input side — shared by all exporters)

We build `utils/html-tokenizer.ts` + `src/html-parser/` instead of these.

| Package | Purpose |
| --- | --- |
| htmlparser2 | Fast streaming HTML tokenizer |
| parse5 | WHATWG HTML parser |
| jsdom | Full DOM implementation |
| cheerio | jQuery-like traversal |
| rehype | HTML AST ecosystem |
| unified | Content processing pipeline |

---

## HTML Entity Decoding (needed by the HTML parser)

HTML input contains entities like `&amp;`, `&nbsp;`, `&#160;`, `&mdash;`.
We decode these inside `utils/html-tokenizer.ts` instead of these.

| Package | Purpose |
| --- | --- |
| he | HTML entity encode/decode (most complete) |
| entities | HTML/XML entity decoding |

---

## ZIP Writer (gates DOCX, XLSX, PPTX output)

We build `utils/zip-writer.ts` instead of these.
Compression uses `node:zlib` `deflateRawSync` — no npm dep needed for DEFLATE.

| Package | Purpose |
| --- | --- |
| jszip | ZIP archive writing (primary reference for write structure) |
| archiver | Streaming ZIP builder |
| adm-zip | Simple ZIP read/write |
| yazl | Lightweight ZIP writer |

---

## XML Builder (gates DOCX, XLSX, PPTX output)

We build `utils/xml-builder.ts` instead of these.

| Package | Purpose |
| --- | --- |
| xmlbuilder2 | Fluent XML builder (primary reference) |
| fast-xml-parser | XML builder/parser reference |

---

## MIME Type Lookup (asset embedding in DOCX, XLSX, PDF)

When embedding images into output files, exporters need the correct MIME type
for `[Content_Types].xml` entries (OOXML) and image dictionaries (PDF).

| Package | Purpose |
| --- | --- |
| mime-types | Extension ↔ MIME type lookup |
| mime | Lightweight MIME type detection |

---

## Image Asset Handling (embed images into output files)

| Package | Purpose |
| --- | --- |
| sharp | Image decode/encode |
| image-size | Read image dimensions from headers |
| file-type | Detect image MIME type from magic bytes |
| jimp | Pure JS image processing |
