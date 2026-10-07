# R&D: How the Reference Exporters Work + Our From-Scratch Plan

**Constraint:** Build every exporter ourselves. We may read the reference packages'
repos / npm pages / source for understanding, but we do **not** take a runtime
dependency on any export library. Node core modules (`zlib`, `buffer`,
`fs`, `crypto`) are fair game — they are not "npm packages."

This document is the R&D output for [AGENT.MD](../AGENT.MD): for each output format it
records *how the reference packages do it* and *how we reproduce it from scratch*,
plus a phased roadmap grounded in the existing scaffold.

---

## 0. The key insight — build primitives, not 8 exporters

The output formats look like 8 unrelated problems. They are not. Almost every
"hard" format reduces to **a small number of shared primitives**. Build these once
and most exporters become thin mapping layers.

| Primitive | Replaces (npm) | Unlocks |
| --- | --- | --- |
| `html-tokenizer.ts` — forgiving HTML tokenizer + tree builder | `htmlparser2`, `parse5`, `cheerio` | **Every exporter** (HTML is the only input) |
| `zip-writer.ts` — ZIP archive builder (`node:zlib` deflate) | `jszip`, `archiver` | **DOCX, XLSX, PPTX** (all are ZIP files) |
| `xml-builder.ts` — XML string serializer | `xmlbuilder2`, `fast-xml-parser` (write) | **DOCX, XLSX, PPTX** (all are XML inside ZIP) |
| `bytes.ts` — binary writer (u16/u32 LE/BE) ✅ already built | — | **PDF** (binary page-description format) |
| `node:zlib` deflate | — | ZIP compression (OOXML) + PDF `FlateDecode` streams |

> **Takeaway:** `html-tokenizer` is the gate for everything (no input without it).
> `zip-writer` + `xml-builder` together unlock the entire Office trio (DOCX/XLSX/PPTX).
> That is the highest-leverage work after the HTML parser.

Everything flows from the **IDM** (`src/idm/index.ts`). The HTML parser fills the
IDM; every exporter reads from it. No exporter ever touches raw HTML directly.

---

## 1. HTML Parser — `htmlparser2` / `parse5` / `cheerio`

This is the **only input path**. Every exporter depends on it.

**How the reference packages work:**

- `htmlparser2` — fast, forgiving, streaming tokenizer. Emits open/close/text/comment
  events. Does not build a DOM itself — you do that on top.
- `parse5` — implements the full WHATWG HTML5 tree-construction algorithm (insertion
  modes, foster-parenting, implicit tag closing). Correct but complex.
- `cheerio` — jQuery-style API over an `htmlparser2` DOM. Convenient for traversal.

**From scratch (pragmatic subset):**

We do not need the full WHATWG spec. Real HTML documents from editors (Notion, Quill,
TipTap) and AI output are well-structured. We need:

1. **Tokenizer** — character state machine emitting tokens:
   - Open tag + attributes: `<p class="intro">`
   - Self-closing / void tags: `<br>`, `<img>`, `<hr>`, `<input>`
   - Close tag: `</p>`
   - Text content
   - Comments: `<!-- ... -->`
   - Raw-text elements: `<script>` and `<style>` content treated as opaque text

2. **Tree builder** — turns the token stream into a DOM tree with auto-close rules:
   - `<p>` auto-closes before another `<p>`, `<h1>`–`<h6>`, `<ul>`, `<table>`
   - `<li>` auto-closes before another `<li>` or `</ul>`/`</ol>`
   - `<td>`/`<th>` auto-close before another `<td>`/`<th>` or `</tr>`
   - `<tr>` auto-closes before another `<tr>` or `</tbody>`

3. **DOM → IDM mapper** — walks the DOM tree and converts each element to its IDM node.
   Drops noise: `<script>`, `<style>`, `<nav>`, `<iframe>`, inline event handlers.

**Feasibility:** Medium. A sane subset handles all real editor output. Full WHATWG
edge cases (misnested tables, `<template>`, foster-parenting) are out of scope.

---

## 2. ZIP Writer — foundation of DOCX / XLSX / PPTX

DOCX, XLSX, and PPTX are ZIP archives. To **generate** them we need to **write** ZIP files.

**How `jszip` / `archiver` work:**

1. Collect all entries (name → bytes)
2. For each entry: write a Local File Header (`PK\x03\x04`) + compressed data
3. After all entries: write the Central Directory (one `PK\x01\x02` record per entry)
4. Write the End-of-Central-Directory record (`PK\x05\x06`)

**From scratch (feasible):**

Compression method `8` = DEFLATE → `zlib.deflateRawSync` (Node core).
Method `0` = stored (no compression) — used for small XML files.

```
zip-writer.ts

  addFile(name, buffer, compress?)
  addFile(name, buffer, compress?)
  ...
  finalize() → Buffer   ← complete valid ZIP
```

The existing `zip.ts` (reader) is the reference — we already know the exact byte
layout of every ZIP structure. Writing is the mirror of reading.

**Acceptance:** write a ZIP → re-read with existing `ZipArchive.read()` → all entries
round-trip correctly.

---

## 3. XML Builder — generates XML for OOXML content

Every file inside a DOCX/XLSX/PPTX is XML. We need to generate that XML cleanly.

**How `xmlbuilder2` / `fast-xml-parser` (write mode) work:**

Fluent builder API that tracks open elements and serializes to a string:

```ts
builder.el('w:document')
  .el('w:body')
    .el('w:p')
      .el('w:r')
        .el('w:t').text('Hello').up()
      .up()
    .up()
  .up()
.end()
// → '<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'
```

**From scratch (easy):**

A simple class that tracks a stack of open elements and builds a string:

```ts
class XmlBuilder {
  el(name: string, attrs?: Record<string, string>): this  // open element
  text(value: string): this                               // text content (escaped)
  up(): this                                              // close current element
  end(): string                                           // return final XML string
}
```

Attribute escaping (`&` → `&amp;`, `<` → `&lt;`, `"` → `&quot;`) reuses the
existing `utils/html.ts` `escapeAttribute` helper.

**Acceptance:** serialize → parse with existing `parseXml()` → structure is identical.

---

## 4. DOCX Exporter — `docx` (npm) / `officegen`

**How the reference packages work:**

They build the WordprocessingML XML tree programmatically and package it into a ZIP.
`docx` (npm) has a TypeScript API that maps to `<w:p>`, `<w:r>`, `<w:tbl>` etc.

**From scratch — what we generate:**

```
word/document.xml       ← paragraphs, headings, tables, images, links
word/styles.xml         ← Heading1–6, Normal, ListBullet, ListNumber styles
word/numbering.xml      ← list numbering definitions (for <ul>/<ol>)
word/_rels/document.xml.rels ← relationships: images, hyperlinks
word/media/image1.png   ← embedded image bytes
docProps/core.xml       ← title, author, created date
[Content_Types].xml     ← MIME types for each file in the ZIP
_rels/.rels             ← top-level relationship to document.xml
```

**IDM → WordprocessingML mapping:**

| IDM Node | XML Output |
| --- | --- |
| `ParagraphNode` | `<w:p><w:r><w:t>text</w:t></w:r></w:p>` |
| `HeadingNode { level: 1 }` | `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>…</w:p>` |
| `StrongNode` | `<w:rPr><w:b/></w:rPr>` inside `<w:r>` |
| `EmphasisNode` | `<w:rPr><w:i/></w:rPr>` inside `<w:r>` |
| `ListNode (unordered)` | `<w:pPr><w:pStyle w:val="ListBullet"/><w:numPr>…</w:numPr></w:pPr>` |
| `ListNode (ordered)` | `<w:pPr><w:pStyle w:val="ListNumber"/><w:numPr>…</w:numPr></w:pPr>` |
| `TableNode` | `<w:tbl><w:tr><w:tc><w:p>…</w:p></w:tc></w:tr></w:tbl>` |
| `LinkNode` | `<w:hyperlink r:id="rId1">` + relationship entry |
| `ImageNode` | `<w:drawing><wp:inline><a:blip r:embed="rId2"/>…</wp:inline></w:drawing>` + media file |

**Feasibility:** High. IDM maps almost 1:1 to WordprocessingML. The trickiest part
is the list numbering XML (`word/numbering.xml`) and the relationship file pattern.

---

## 5. XLSX Exporter — `exceljs` / `xlsx` (SheetJS write mode)

**How the reference packages work:**

Build SpreadsheetML XML for each worksheet, collect cell values into a shared string
table, package into ZIP.

**From scratch — what we generate:**

```
xl/workbook.xml             ← sheet list + active sheet
xl/worksheets/sheet1.xml    ← grid of rows and cells
xl/sharedStrings.xml        ← deduplicated string table
xl/styles.xml               ← cell styles (bold header, date format, etc.)
xl/_rels/workbook.xml.rels  ← relationships to sheet files
[Content_Types].xml
_rels/.rels
```

**IDM → SpreadsheetML mapping:**

Each `TableNode` becomes one worksheet. Cells use A1 notation (`r="A1"`, `r="B2"`):

```xml
<sheetData>
  <row r="1">
    <c r="A1" t="s" s="1"><v>0</v></c>   <!-- t="s"=shared string, s="1"=bold style -->
    <c r="B1" t="s" s="1"><v>1</v></c>
  </row>
  <row r="2">
    <c r="A2" t="s"><v>2</v></c>
    <c r="B2" t="s"><v>3</v></c>
  </row>
</sheetData>
```

Shared strings: all cell text goes into `xl/sharedStrings.xml` and cells reference
by index. This deduplicates repeated values and is mandatory for large sheets.

**Feasibility:** High. Fiddly parts are A1 reference math (column letter from index),
shared strings indexing, and the styles XML for header bold formatting.

---

## 6. PPTX Exporter — `pptxgenjs`

**How `pptxgenjs` works:** builds PresentationML XML per slide, packages into ZIP.
Each slide is a separate XML file. Slide order defined in `presentation.xml` via rels.

**From scratch — what we generate:**

```
ppt/presentation.xml              ← slide list + slide size
ppt/slides/slide1.xml             ← content of slide 1
ppt/slides/_rels/slide1.xml.rels  ← slide relationships (images)
ppt/slideLayouts/slideLayout1.xml ← layout (inherited formatting)
ppt/slideMasters/slideMaster1.xml ← master (theme)
ppt/media/image1.png              ← embedded images
ppt/notesSlides/notesSlide1.xml   ← speaker notes
```

**IDM → PresentationML mapping:**

Each `SectionNode` → one slide. Inside a slide:

```xml
<p:sp>
  <p:nvSpPr>
    <p:nvPr><p:ph type="title"/></p:nvPr>   ← title placeholder
  </p:nvSpPr>
  <p:txBody>
    <a:p><a:r><a:t>Slide Title</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
```

Body text, bullet lists → `<p:ph type="body"/>` shape with `<a:pPr lvl="N">` for
nesting level and `<a:buChar char="•"/>` for bullet character.

**Feasibility:** High — reuses zip-writer + xml-builder; mostly different tag vocabulary
from DOCX/XLSX. The boilerplate (slide master/layout XML) is the most tedious part.

---

## 7. Markdown Exporter — `turndown` / `node-html-markdown`

**How they work:** walk the HTML DOM (or our IDM) and emit Markdown syntax for each
element. The hard parts are nested lists (indentation tracking) and GFM tables.

**From scratch (easy):**

Since we work from the IDM (not raw HTML), every node type has a known, clean
Markdown equivalent. No ambiguity about what `<b>` vs `<strong>` means — it is
always `StrongNode` by the time the exporter sees it.

```
HeadingNode { level }   → "#".repeat(level) + " " + text + "\n\n"
ParagraphNode           → text + "\n\n"
StrongNode              → "**" + text + "**"
EmphasisNode            → "*" + text + "*"
CodeSpanNode            → "`" + text + "`"
LinkNode                → "[" + text + "](" + href + ")"
ListNode (unordered)    → "- " per item (indent 2 spaces per nesting level)
ListNode (ordered)      → "1. " per item
TableNode               → pipe table: "| col | col |\n|---|---|\n| val | val |"
CodeNode                → "```" + lang + "\n" + code + "\n```"
QuoteNode               → "> " prefix on every line
ImageNode               → "![" + alt + "](" + src + ")"
ThematicBreakNode       → "---"
```

**Feasibility:** Easy. The IDM already encodes all the information; this is pure
text serialization.

---

## 8. CSV Exporter — `json2csv` / `fast-csv` (write mode)

**How they work:** iterate rows, RFC 4180 quote fields that contain commas/newlines/
double-quotes, join with `\r\n`.

**From scratch (trivial):**

The existing `parseCsv()` in `src/parser/csv/index.ts` is the read-side reference.
The write side is simpler — just the quoting rule:

```
field contains , or " or \n  → wrap in "..." and escape internal " as ""
otherwise                    → write as-is
```

One `TableNode` → one CSV output. Multiple tables → multiple named assets.

**Feasibility:** Trivial. Already have the RFC 4180 state machine for reading;
writing is simpler.

---

## 9. TXT Exporter — `html-to-text`

**How it works:** strip HTML, preserve structure with whitespace conventions.
Headings become underlined or ALL CAPS, lists get `-` prefixes, tables become
ASCII art grids.

**From scratch (easy):**

```
HeadingNode { level: 1 } → "HEADING\n" + "=".repeat(len) + "\n\n"
HeadingNode { level: 2 } → "Heading\n" + "-".repeat(len) + "\n\n"
ParagraphNode            → text + "\n\n"
ListNode (unordered)     → "- item\n" (2-space indent per level)
ListNode (ordered)       → "1. item\n"
TableNode                → ASCII art: "| col | col |\n|-----|-----|\n| val | val |"
LinkNode                 → "text (url)"
ImageNode                → "[image: alt]"
```

Column widths for ASCII tables: scan all cells in each column, use the max length.

**Feasibility:** Easy.

---

## 10. JSON Exporter

**How it works:** serialize the IDM tree to a JSON document. Tables get special
treatment — rendered as `{ headers, rows }` for immediate programmatic use.

**From scratch (trivial):**

The IDM is already a plain JavaScript object tree. The exporter is essentially
`JSON.stringify` with a few transformations (table flattening, metadata lifting).

**Feasibility:** Trivial.

---

## 11. PDF Exporter — `pdfkit` / `pdf-lib` / `jsPDF` *(the hard one)*

**How the reference packages work:**

PDF is a page-description language, not a document format. You specify exact
positions of text and graphics:

```
%PDF-1.4
<indirect objects: catalog, pages, fonts, images, content streams>
xref table
trailer
%%EOF
```

Content streams use operators:
```
BT                    % begin text
/Helvetica 12 Tf      % select font + size
50 750 Td             % move to (x, y) — origin is bottom-left
(Hello World) Tj      % show string
ET                    % end text
```

Page layout requires tracking the Y cursor manually. When Y < bottom margin: new page.

**From scratch — scope in tiers:**

1. **Skeleton:** `%PDF-1.4` header, indirect objects, `xref` table, `trailer` dict, minimal catalog + page tree. Uses `bytes.ts` for binary writing.
2. **Text:** embed Helvetica (standard Type 1 font, no embedding needed); paragraph flow with Y-cursor tracking; line wrapping by character count.
3. **Headings:** font size variants (24pt H1, 18pt H2, 14pt H3, 12pt body); bold via `/Helvetica-Bold`.
4. **Page breaks:** new page when Y overflows; `PageBreakNode` triggers explicit break.
5. **Images:** JPEG → `DCTDecode` stream (raw bytes); PNG → decode + `FlateDecode`.
6. **Tables:** cell borders via `re S` path operators; text inside cells at calculated positions.
7. **Hyperlinks:** URI action annotations at exact bounding boxes of link text.
8. **Metadata:** `Info` dict (Title, Author, CreationDate) in trailer.

**Feasibility:** Low-Medium — highest risk/effort of all exporters. Build last,
ship in tiers so basic text+headings lands well before tables and images.

---

## 12. Image Export — `puppeteer` / `playwright` ⚠️ decision needed

**Reality:** rendering HTML to a pixel-accurate PNG/JPEG requires a layout engine
(CSS box model, font metrics, image compositing) plus a rasterizer. This is
essentially a browser.

Building a layout engine from scratch is **not realistic** within this project.

**Options to decide:**

- **(A) Defer** entirely. Image export is not in Phase 1–4. *(recommended)*
- **(B)** Allow `puppeteer` or `playwright` as a single carve-out from the no-npm rule — it is an engine (like the OCR decision was for the old direction), not a parsing library.
- **(C)** Ship a deliberately simple rasterizer (monospace text only, no CSS) — educational only.

---

## 13. Cross-cutting: performance / streaming

- **XLSX large sheets:** write rows to the worksheet XML stream incrementally rather
  than building the full XML in memory. `xl/sharedStrings.xml` can be flushed at the
  end once all strings are collected.
- **PDF large documents:** write each page's content stream independently; append
  to the output buffer page-by-page rather than building the full document in memory.
- **ZIP writer:** compress each entry independently; no need to hold all entries
  in memory simultaneously if using streaming.
- **HTML parser:** process the input as a string (in-memory); for very large HTML
  files (>10 MB), consider a streaming tokenizer that avoids building the full DOM.

---

## 14. Phased roadmap

> **Granular per-exporter milestones live in [development-phases.md](development-phases.md).**
> This table is the high-level overview.

| Phase | Deliverable | Risk | Notes |
| --- | --- | --- | --- |
| **0 — New primitives** | `html-tokenizer`, `zip-writer`, `xml-builder` | Low | Gates everything else |
| **1 — HTML Parser** | HTML string → IDM `DocumentNode` | Med | Critical path — all exporters depend on it |
| **2 — Text exporters** | Markdown, CSV, TXT, JSON | Low | Validate pipeline end-to-end; no binary complexity |
| **3 — OOXML exporters** | DOCX → XLSX → PPTX | Low-Med | Pure ZIP+XML generation; reuses Phase 0 primitives |
| **4 — PDF** | Tiers 1–4 (text → tables → images → links) | High | Biggest effort; ship incrementally |
| **5 — Deferred** | Image export (PNG/JPEG), CLI, `convertAll()` | High | Optional / engine-gated |

---

## 15. Immediate next steps

Build the **three new Phase 0 primitives** in this order:

1. **`html-tokenizer.ts`** — gates Phase 1 (HTML parser), which gates everything
2. **`zip-writer.ts`** — gates Phase 3 (DOCX/XLSX/PPTX); prove it with a round-trip test against existing `zip.ts` reader
3. **`xml-builder.ts`** — gates Phase 3; prove it with a round-trip test against existing `parseXml()`

Then move to Phase 1 (HTML parser) immediately — without a valid IDM from real HTML,
no exporter can be tested against real input.
