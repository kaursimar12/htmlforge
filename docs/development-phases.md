# Development Phases — HTML → Any File Exporter

This tracker breaks the full build into incremental phases. Every phase ships
thin, testable slices. No phase depends on a later one.

**Legend:** ✅ done · 🔜 next up · ⬜ planned

---

## Phase 0 — Shared Primitives

The toolbox everything else depends on. Build these before touching any exporter.

### Already built ✅

| ID | Module | Used by |
| --- | --- | --- |
| P0.1 | `utils/bytes.ts` — binary reader/writer | PDF exporter |
| P0.2 | `utils/xml.ts` — XML parser + traversal | Reference for xml-builder |
| P0.3 | `utils/zip.ts` — ZIP reader | Reference for zip-writer |
| P0.4 | `utils/magic.ts` — magic-number detection | Detect embedded image types in assets |

### To build 🔜

| ID | Module | Replaces (npm) | Needed by | Status |
| --- | --- | --- | --- | --- |
| P0.5 | `utils/html-tokenizer.ts` — forgiving HTML tokenizer + tree builder | `htmlparser2`, `parse5` | Phase 1 HTML Parser | 🔜 |
| P0.6 | `utils/zip-writer.ts` — ZIP archive builder (`node:zlib` deflate) | `jszip`, `archiver` | DOCX, XLSX, PPTX exporters | 🔜 |
| P0.7 | `utils/xml-builder.ts` — XML string serializer | `xmlbuilder2` | DOCX, XLSX, PPTX exporters | 🔜 |

**Acceptance — P0.5:** tokenize a fragment of messy HTML with unclosed tags and void elements; verify all tokens are captured correctly.

**Acceptance — P0.6:** write a ZIP with 3 entries → re-read with existing `ZipArchive.read()` → all entries round-trip byte-for-byte.

**Acceptance — P0.7:** serialize an XML tree with namespace prefixes → re-parse with existing `parseXml()` → structure is identical.

---

## Phase 1 — HTML Parser (Critical Path) 🔜

The only input to the entire system. Every single exporter depends on this phase.
Without it, nothing can be tested against real HTML.

**Flow:** HTML string → tokenizer → DOM tree → IDM `DocumentNode`

| # | Milestone | Deps | Status |
| --- | --- | --- | --- |
| H1 | `html-parser/tokenizer.ts` — emit tokens for open tags, close tags, attributes, text, comments, void elements (`br`, `img`, `hr`), raw-text elements (`script`, `style`) | P0.5 | 🔜 |
| H2 | `html-parser/tree-builder.ts` — assemble DOM tree from token stream; auto-close rules: `<p>` before block elements, `<li>` before `<li>`, `<td>`/`<th>` before next cell, `<tr>` before next row | H1 | ⬜ |
| H3 | `html-parser/dom-to-idm.ts` — map DOM nodes to IDM: `h1–h6` → `HeadingNode`, `p` → `ParagraphNode`, `ul/ol/li` → `ListNode/ListItemNode`, `table/thead/tbody/tr/td/th` → `TableNode/TableRowNode/TableCellNode`, `a` → `LinkNode`, `img` → `ImageNode`, `figure/figcaption` → `FigureNode`, `strong/b` → `StrongNode`, `em/i` → `EmphasisNode`, `code` → `CodeSpanNode`, `pre` → `CodeNode`, `blockquote` → `QuoteNode`, `section/article` → `SectionNode`, `br` → `LineBreakNode`, `hr` → `ThematicBreakNode` | H2 | ⬜ |
| H4 | Strip noise during DOM→IDM walk: silently drop `script`, `style`, `nav`, `iframe`, `noscript`, inline event handlers (`onclick`, `onload`, etc.) | H3 | ⬜ |
| H5 | Metadata extraction from `<head>`: `<title>` → `Metadata.title`; `<meta name="author">` → `Metadata.author`; `<html lang="en">` → `Metadata.language`; `<meta name="description">` → `Metadata.description` | H3 | ⬜ |
| H6 | Asset resolver: fetch image bytes from `http://` URLs (via `node:https`), `data:` URIs (base64 decode), local file paths (via `node:fs`); push resolved bytes as `Asset` records; set `ImageNode.assetId` | H3 | ⬜ |

**Acceptance:** Parse a real-world HTML page (a Wikipedia article or a Notion export) → IDM has correct heading hierarchy, all paragraphs, lists, tables, and resolved image assets. Run against 10 real documents; zero crashes.

---

## Phase 2 — Text Exporters

Pure text output — no ZIP, no binary structures. These four exporters ship first
because they validate the complete pipeline (HTML → IDM → output) with zero
binary complexity. Any IDM design issues surface here before we touch OOXML.

---

### 2A — Markdown Exporter ⬜

Reference npm packages: `turndown`, `node-html-markdown`, `rehype-remark`

| # | Milestone | Consumes (IDM) | Status |
| --- | --- | --- | --- |
| M1 | Headings → `#`–`######`; paragraphs → blank-line separated | `HeadingNode`, `ParagraphNode`, `TextNode` | ⬜ |
| M2 | Inline: `StrongNode` → `**text**`, `EmphasisNode` → `*text*`, `CodeSpanNode` → `` `text` ``, `LinkNode` → `[text](href)` | inline nodes | ⬜ |
| M3 | Lists: `ListNode (unordered)` → `- item`; `ListNode (ordered)` → `1. item`; nested lists → 2-space indent per level | `ListNode`, `ListItemNode` | ⬜ |
| M4 | GFM tables → pipe-delimited rows with `|---|` separator row | `TableNode`, `TableRowNode`, `TableCellNode` | ⬜ |
| M5 | `CodeNode` → fenced ` ``` ` block with language; `QuoteNode` → `> ` prefix on every line | `CodeNode`, `QuoteNode` | ⬜ |
| M6 | `ImageNode` → `![alt](src)`; `FigureNode` → image + `*caption*` on next line | `ImageNode`, `FigureNode` | ⬜ |
| M7 | `ThematicBreakNode` → `---`; `PageBreakNode` → `---` | `ThematicBreakNode`, `PageBreakNode` | ⬜ |

**Acceptance:** Golden snapshot test — an HTML fixture covering all element types → expected `.md` string verified character-for-character.

---

### 2B — CSV Exporter ⬜

Reference npm packages: `json2csv`, `fast-csv`, `html-table-to-json`

| # | Milestone | Consumes (IDM) | Status |
| --- | --- | --- | --- |
| C1 | First `TableNode` → RFC 4180 CSV rows; `<thead>` / first row of `<th>` cells → header row | `TableNode`, `TableRowNode`, `TableCellNode` | ⬜ |
| C2 | RFC 4180 quoting: fields with `,` / `"` / newlines wrapped in `"..."`; internal `"` escaped as `""` | — | ⬜ |
| C3 | Multiple tables → multiple named CSV assets (`table-1.csv`, `table-2.csv`); emit warning when >1 table found | — | ⬜ |

**Acceptance:** HTML with a 3-column table → CSV that re-parses via the existing `parseCsv()` to identical rows.

---

### 2C — Plain Text Exporter ⬜

Reference npm packages: `html-to-text`, `striptags`

| # | Milestone | Consumes (IDM) | Status |
| --- | --- | --- | --- |
| T1 | `HeadingNode (level 1)` → `ALL CAPS + ===== underline`; level 2 → `Title Case + ----- underline`; level 3+ → `Title Case` with blank lines | `HeadingNode` | ⬜ |
| T2 | Paragraphs → blank-line separated; `LineBreakNode` → newline | `ParagraphNode` | ⬜ |
| T3 | Lists: `- item` (unordered), `1. item` (ordered); 2-space indent per nesting level | `ListNode`, `ListItemNode` | ⬜ |
| T4 | Tables → ASCII art grid; column widths auto-sized to max content length per column | `TableNode` | ⬜ |
| T5 | `LinkNode` → `text (url)`; `ImageNode` → `[image: alt]`; `CodeSpanNode` → `` `text` `` | inline nodes | ⬜ |
| T6 | `CodeNode` → 4-space indented block; `QuoteNode` → `> ` prefix | `CodeNode`, `QuoteNode` | ⬜ |

**Acceptance:** Output is human-readable and contains all meaningful content from the input HTML; no HTML tags or entities appear in the output.

---

### 2D — JSON Exporter ⬜

Reference npm packages: `htmlparser2 + domhandler`, `cheerio`

| # | Milestone | Consumes (IDM) | Status |
| --- | --- | --- | --- |
| J1 | IDM → JSON object tree mirroring the node structure: `{ type, children }` for block nodes, `{ type, value }` for text nodes | `DocumentNode`, all block nodes | ⬜ |
| J2 | Tables → `{ type: "table", headers: [...], rows: [[...]] }` — flat array-of-arrays for easy consumption | `TableNode` | ⬜ |
| J3 | Metadata → top-level `meta` key; resolved assets → `assets` array with `{ id, src, mimeType }` | `Metadata`, `Asset` | ⬜ |

**Acceptance:** JSON output `JSON.parse()`s without error; all headings, paragraph text, and table rows present and correct; `JSON.stringify(result)` is idempotent.

---

## Phase 3 — OOXML Exporters

All three Office formats are **ZIP archives of XML files**. They all depend on
`utils/zip-writer.ts` (P0.6) and `utils/xml-builder.ts` (P0.7).

They share a common helper `src/exporter/ooxml/package.ts` that generates
the boilerplate files every OOXML package requires:

```
[Content_Types].xml     ← MIME types for every file in the ZIP
_rels/.rels             ← top-level relationship pointing to the main document
docProps/core.xml       ← title, author, created/modified dates
docProps/app.xml        ← application name
```

Without these exact files at these exact paths, Office refuses to open the file.

---

### 3A — DOCX Exporter ✅

Reference npm packages: `html-to-docx`, `officegen`

A prior attempt at this exporter reached "76 passing tests" and was reverted as
totally wrong — it opened in Word as corrupted/unreadable in several cases
despite green tests. Root causes (recovered from git history — see
`src/exporter/docx/`, `src/exporter/ooxml/package.js`): blockquotes rendered
an empty stub paragraph instead of styling the real content; tables never
emitted `w:tblGrid`/`w:vMerge`, so rowspan/colspan produced mismatched cell
counts per row; images always fell back to `[Image: alt]` placeholder text
because the relationship list was declared but never populated; and every
OOXML test only asserted "starts with PK" / "mimeType contains X" — never
whether the generated XML was actually well-formed or structurally
consistent. The rebuild fixes all four and is covered by 40+ structural tests
that parse every generated part.

ZIP contents we generate:

```
word/document.xml              ← paragraphs, headings, tables, images, links
word/styles.xml                ← Heading1–6, Normal, ListBullet, ListNumber, Quote, Hyperlink, TableGrid styles
word/numbering.xml             ← list numbering definitions
word/_rels/document.xml.rels   ← styles/numbering/image/hyperlink relationships
word/media/image1.png          ← embedded image bytes
```

| # | Milestone | Deps | Consumes (IDM) | Status |
| --- | --- | --- | --- | --- |
| W1 | Skeleton: valid empty `.docx` — all boilerplate XML + ZIP packaging; opens in Word | P0.6, P0.7 | — | ✅ |
| W2 | Paragraphs: `<w:p><w:r><w:t>text</w:t></w:r></w:p>` | W1 | `ParagraphNode`, `TextNode` | ✅ |
| W3 | Headings: `<w:pStyle w:val="Heading1"/>` through `Heading6`; `<hr>` → bottom-bordered paragraph; page-break `<hr>` → `<w:br w:type="page"/>` | W2 | `HeadingNode`, `ThematicBreakNode`, `PageBreakNode` | ✅ |
| W4 | Inline formatting: `<w:b/>` bold, `<w:i/>` italic (accumulated through nesting), `<w:rFonts w:ascii="Courier New"/>` + shading for code; blockquote paragraphs get the real `Quote` `w:pStyle` directly (no empty stub — the exact bug from the prior attempt); multi-line code block joins lines with `w:br` in one shaded paragraph | W2 | `StrongNode`, `EmphasisNode`, `CodeSpanNode`, `QuoteNode`, `CodeNode` | ✅ |
| W5 | Lists: two shared `abstractNum` definitions (one bullet cycling •/◦/▪, one decimal-family cycling decimal/lowerLetter/lowerRoman across 9 levels); every distinct `<ul>`/`<ol>` — including a nested one — registers its own `numId` via `NumberingManager` so separate lists number independently; nesting depth → `w:ilvl`; `<ol start="N">` → `w:lvlOverride`/`w:startOverride` | W2 | `ListNode`, `ListItemNode` | ✅ |
| W6 | Tables: the IDM's collapsed rowspan/colspan grid is expanded back into a full OOXML grid — `<w:tblGrid>` gets one `<w:gridCol>` per real column, rows a rowspan covers get a placeholder `<w:tc>` with a bare `<w:vMerge/>` matching the origin's `gridSpan` (not just `<w:tbl><w:tr><w:tc>` with no grid, which is what corrupted the prior attempt's tables); header row/cells bold via `w:b`; a cell holding multiple paragraphs or a nested `<table>` renders as multiple `<w:p>`/a nested `<w:tbl>` inside one `<w:tc>` (OOXML natively supports this, unlike the ASCII/GFM exporters) | W2 | `TableNode`, `TableRowNode`, `TableCellNode` | ✅ |
| W7 | Hyperlinks: External relationship entry in `document.xml.rels`; `<w:hyperlink r:id="rId…">` wraps runs styled with the `Hyperlink` character style | W3 | `LinkNode` | ✅ |
| W8 | Images: `resolveAssets()` is called with `embedImages` forced on (a non-embedded reference makes no sense in a self-contained `.docx`); real bytes go to `word/media/`, a real `image`-type relationship is registered, and `<w:drawing><wp:inline>…<a:blip r:embed="rId…"/>` is emitted — not a `[Image: alt]` text placeholder, which is what the prior attempt always fell back to since its relationship list was populated in name only | W3 | `ImageNode`, `FigureNode`, `Asset` | ✅ |
| W9 | Metadata: title/author/language → `docProps/core.xml` (`document.metadata`, or `convert()`'s `title`/`author`/`language` options, which override); nothing is fabricated when absent (no empty `dc:title`, no synthesized `created` date) | W1 | `Metadata` | ✅ |
| W10 | Page setup: `convert()`'s `pageSize` (`"Letter"`/`"A4"`/`"Legal"`)/`orientation`/`margins` (mm) options → `<w:sectPr>`'s `w:pgSz`/`w:pgMar`, defaulting to US Letter portrait with 1in margins | W1 | — (reads `ExporterContext.options`) | ✅ |

**Acceptance:** `tests/exporter-docx.test.js` (40+ tests) parses every generated
part with the project's own `zip.js`/`xml.js` readers and asserts structural
conformance per milestone — not just "starts with PK." A whole-document
integration test (headings, nested lists, a rowspan/colspan table, code,
blockquote, a figure with a real embedded image, and a link, all in one
fixture) verifies every `r:id`/`r:embed` referenced in `document.xml`
resolves in `.rels` and every media relationship's target file actually
exists in the zip — the exact class of check the prior attempt never had at
any granularity. **Still required, and not yet done in this environment**
(no LibreOffice/Word available here): open a generated `.docx` in real Word
or LibreOffice Writer via `node dev/preview-server.mjs` and visually confirm
headings, nested lists, the rowspan/colspan table, blockquote styling, the
embedded image with caption, and the page break all render correctly.

---

### 3B — XLSX Exporter ✅

Reference npm packages: `node-html-to-xlsx`, `exceljs`, `xlsx (SheetJS)`

The original scope below (X1–X6) only asked for plain values and bold
headers. After reviewing a real-world target document — a Google-Sheets-style
invoice template with a merged/colored header banner, banded row fills,
section-divider borders, right-aligned totals, currency (`$50.00`) and
percent (`5.00%`) formatting, and a merged footer bar — the scope was
expanded to fold merged cells, cell fills/borders/alignment, and
currency/percent number formats into this same pass rather than deferring
them, since the IDM's `TableCellNode` already carries the raw ingredients
(`colspan`/`rowspan`/`backgroundColor`/`border*`/`align`) via the work
already done for the DOCX exporter. This adds milestone X7 and broadens X2/X3
beyond their original one-line description; see the full design plan at
`C:\Users\Hp\.claude\plans\now-lets-go-to-agile-church.md` for the detailed
cell-value-classification rules, the composite cell-style registry design,
and the merge-cell layout algorithm.

ZIP contents we generate:

```
xl/workbook.xml              ← sheet list
xl/worksheets/sheet1.xml     ← grid of rows + cells, incl. <cols>/<mergeCells>
xl/sharedStrings.xml         ← deduplicated string table
xl/styles.xml                ← cell styles: fonts, fills, borders, number formats, cellXfs
xl/_rels/workbook.xml.rels   ← sheet file relationships
```

| # | Milestone | Deps | Consumes (IDM) | Status |
| --- | --- | --- | --- | --- |
| X1 | Skeleton: valid empty `.xlsx` — all boilerplate XML + ZIP packaging; opens in Excel | P0.6, P0.7 | — | ✅ |
| X2 | Each `TableNode` → one worksheet; cells use A1 notation (`r="A1"`, `r="B2"`); full value classification — plain numbers, percent (`5.00%`), currency (`$`/`€`/`£`), and thousands-separated (`1,234`) all render as real numeric cells with a matching number format; everything else (incl. empty cells, which are omitted) goes through shared strings | X1 | `TableNode`, `TableRowNode`, `TableCellNode` | ✅ |
| X3 | Full cell styling via `xl/styles.xml`'s composite registry: bold (`<thead>`/`<th>` and explicit `font-weight`), font color, background fill, per-side borders, alignment, and the number formats from X2 all apply through deduplicated font/fill/border/numFmt/cellXf entries | X2 | `TableCellNode` (`header`, `bold`, `color`, `backgroundColor`, `border*`, `align`) | ✅ |
| X4 | Shared strings: all cell text → `xl/sharedStrings.xml`; cells use `t="s"` with index; dedup verified across sheets (one shared table for the whole workbook) | X2 | `TextNode` | ✅ |
| X5 | Multiple tables → multiple sheets; sheet names from `<caption>` or nearest heading, sanitized/truncated to 31 chars/de-duplicated, else auto `Sheet1`, `Sheet2` | X2 | `TableNode`, `SectionNode` | ✅ |
| X6 | Column auto-width: scan max content length per column; set `<col width="..."/>` | X3 | — | ✅ |
| X7 | Merged cells: `colspan`/`rowspan` → `<mergeCells>`, via an occupancy-tracking grid layout (conceptually like DOCX's `expandTableGrid`, adapted for Excel's simpler "no continuation cell" merge model) | X2 | `TableCellNode` (`colspan`, `rowspan`) | ✅ |

**Acceptance:** `tests/exporter-xlsx.test.js` (37 tests) parses every
generated part with the project's own `zip.js`/`xml.js` readers and asserts
structural conformance per milestone, including the pure `cell-value.js`
classification rules in isolation, a merged-cell/multi-sheet/dedup-checking
suite, and a whole-document acceptance fixture modeled on the real invoice
template that motivated X7 and the X2/X3 scope expansion (merged colored
header banner, currency/percent values, right-aligned totals, merged footer
bar) — verifying correct merges, styles, and numeric values by A1 ref.
**Still required, and not yet done in this environment** (no Excel/
LibreOffice available here): open a generated `.xlsx` in real Excel or
LibreOffice Calc via `node dev/preview-server.mjs` and visually confirm the
merges, fills/borders, and number formats render as expected — the same
caveat the DOCX exporter's test suite carries.

New shared module: `src/idm/flatten-text.js` (IDM → plain text, used by both
CSV and XLSX for cell values, and by DOCX for its content-based column-width
heuristic) — pulled out of the CSV exporter once XLSX needed the identical
logic. `src/exporter/ooxml/package.js`'s `contentTypesXml`/`rootRelsXml` were
generalized (an `overrides` param and a `mainPart` param, respectively) so
XLSX could reuse them instead of duplicating docx's hardcoded `/word/*`
assumptions — `NS`, `RelationshipManager`, `corePropsXml`, `appPropsXml` were
already fully reusable as-is.

---

### 3C — PPTX Exporter ⬜

Reference npm packages: `pptxgenjs`, `officegen`

ZIP contents we generate:

```
ppt/presentation.xml               ← slide list + slide size
ppt/slides/slide1.xml              ← content of slide 1
ppt/slides/_rels/slide1.xml.rels   ← slide image relationships
ppt/slideLayouts/slideLayout1.xml  ← layout (inherited formatting)
ppt/slideMasters/slideMaster1.xml  ← master theme
ppt/media/image1.png               ← embedded images
ppt/notesSlides/notesSlide1.xml    ← speaker notes
```

| # | Milestone | Deps | Consumes (IDM) | Status |
| --- | --- | --- | --- | --- |
| PP1 | Skeleton: valid empty `.pptx` — presentation + one slide + master + layout boilerplate; opens in PowerPoint | P0.6, P0.7 | — | ⬜ |
| PP2 | Each `SectionNode` → one slide; slide order in `ppt/presentation.xml` rels | PP1 | `SectionNode` | ⬜ |
| PP3 | `HeadingNode (level 1/2)` → title placeholder `<p:ph type="title"/>`; other paragraphs → body text box `<p:ph type="body"/>` | PP2 | `HeadingNode`, `ParagraphNode` | ⬜ |
| PP4 | Lists → bullet `<a:buChar char="•"/>` / numbered `<a:buAutoNum type="arabicPeriod"/>` inside `<p:txBody>`; nesting via `<a:pPr lvl="N"/>` | PP2 | `ListNode`, `ListItemNode` | ⬜ |
| PP5 | Images → picture shape `<p:pic>`; bytes in `ppt/media/`; position + size in EMUs (914400 per inch); auto-fit to slide area | PP2 | `ImageNode`, `FigureNode`, `Asset` | ⬜ |
| PP6 | Tables → graphic frame `<p:graphicFrame>` / `<a:tbl>` / `<a:tr>` / `<a:tc>` | PP2 | `TableNode` | ⬜ |
| PP7 | Speaker notes: `FootnoteNode` content → `ppt/notesSlides/notesSlideN.xml` | PP2 | `FootnoteNode` | ⬜ |

**Acceptance:** Generated `.pptx` opens in PowerPoint and LibreOffice Impress; 3-section fixture renders as 3 slides with titles, bullet lists, and an embedded image.

---

## Phase 4 — PDF Exporter ✅

PDF is a **page-description language** — the hardest output format. Text and
graphics are placed at exact X/Y coordinates on a page. The main challenge is
**text flow**: tracking the Y cursor, wrapping lines, and starting new pages.

Reference npm packages: `pdfkit`, `pdf-lib`, `jsPDF`

Built from scratch on `node:zlib` (`FlateDecode` for content streams and
supported-tier PNGs) and this project's own `ByteWriter` — no PDF library
dependency. The one genuinely new capability versus every other exporter is
real pagination: `PdfPageFlow.ensureSpace(heightPt)` is called at the finest
natural per-block unit (line/row/item), so a block taller than a page is a
non-event rather than a special case, and mixed bold/italic/code inline runs
are wrapped by flattening to leaf runs, computing line breaks via
`estimateWrappedLines`, then re-deriving per-line per-style segments from a
flat word queue. Text uses the standard 14 fonts (`Helvetica`/`Times`/
`Courier` families, no embedding) under `WinAnsiEncoding`; a small
codepoint→byte table (`encodeWinAnsi`) covers common punctuation outside
Latin-1 (smart quotes, en/em dash, bullet, ellipsis) that would otherwise
silently truncate to the wrong byte. Images embed via JPEG `DCTDecode`
passthrough or, for non-interlaced 8-bit grayscale/truecolor PNGs, a
`FlateDecode`+`Predictor 15` passthrough of the concatenated `IDAT` chunks
(no real pixel decode); unsupported PNG tiers (palette, alpha, 16-bit,
interlaced) skip with a placeholder and a `pdf-unsupported-png` warning.
Structure verified by a hand-written PDF reader (`tests/helpers/pdf.js`) that
parses the real xref/trailer/object graph and content-stream operators back
out — not just "starts with `%PDF-`" — the same "verify structure, not
surface bytes" lesson learned from the DOCX rebuild.

ZIP-equivalent (single-file, not a ZIP) contents we generate:

```
%PDF-1.4 header + binary marker
Catalog → Pages → Page objects (one per page, MediaBox + Resources + Contents)
Font objects (only the standard-14 base fonts actually used)
Image XObjects (DCTDecode / FlateDecode+Predictor15, deduped by asset)
Content stream per page (FlateDecode), Link annotations, Info dict
Classic table xref + trailer + startxref/%%EOF
```

| # | Tier | Deps | Consumes (IDM) | Status |
| --- | --- | --- | --- | --- |
| PDF1 | Skeleton: `%PDF-1.4` header, indirect objects (`N G obj … endobj`), classic table `xref`, `trailer` dict, Catalog → Pages → Page tree; object numbers allocated in one fixed pass after the full document walk | P0.1 | — | ✅ |
| PDF2 | Text: standard Helvetica font (no embedding); `BT /F1 11 Tf 1 0 0 1 x y Tm (text) Tj ET` content stream; `cursorY` tracking; word-wrap via `estimateWrappedLines`, breaking across pages when content overflows | PDF1 | `ParagraphNode`, `TextNode` | ✅ |
| PDF3 | Headings: size tiers `{1:24, 2:18, 3:14, 4:12, 5:11, 6:10}`pt, always bold via `/Helvetica-Bold` | PDF2 | `HeadingNode` | ✅ |
| PDF4 | Inline formatting: bold → `Helvetica-Bold`, italic → `Helvetica-Oblique`, code span → `Courier` (always wins regardless of surrounding bold/italic); mixed runs re-derived per-line from a flattened word queue so `Tf` only switches on an actual style change | PDF2 | `StrongNode`, `EmphasisNode`, `CodeSpanNode` | ✅ |
| PDF5 | Lists: `•` bullet (unordered) or `start`-aware decimal marker (ordered), fixed 18pt indent step per nesting depth, drawn item-by-item with per-line `ensureSpace` | PDF2 | `ListNode`, `ListItemNode` | ✅ |
| PDF6 | Page breaks: automatic when `cursorY` would go below the bottom margin; explicit `PageBreakNode` forces a new page unconditionally, even with room left | PDF2 | `PageBreakNode` | ✅ |
| PDF7 | Images: JPEG → `DCTDecode` passthrough (SOF-marker parsed for width/height/components); PNG → `FlateDecode`+`Predictor 15` passthrough for the supported tier, skip+warning otherwise; aspect-fit scaled to content width, capped to one page's height | PDF2 | `ImageNode`, `FigureNode`, `Asset` | ✅ |
| PDF8 | Tables: `expandTableGrid` → content-based column widths (`computeContentBasedColumnWidths` — estimates each column's real text-width need via `estimateTextWidth` and distributes space proportionally, the same approximation of a browser's own table auto-layout DOCX's `computeContentBasedWidths` already does; an explicit `width="%"` on every first-row cell wins over estimation), per-cell wrapped text, row-by-row drawing with header-row redraw on a mid-table page break; a single row taller than one page clips silently (same accepted precedent as PPTX's own `bottomY` clip); a header cell's own explicit CSS background (`th{background:#2563eb}`) always wins over the flat gray default, which only applies when the cell declared none of its own | PDF2 | `TableNode` | ✅ |
| PDF9 | Hyperlinks: `/Annot /Subtype /Link` with `/Rect` computed from the drawn glyph run plus `/A <</S /URI /URI (href)>>`; an internal `#anchor` href renders as plain text with no annotation (no named-destination navigation this pass) | PDF2 | `LinkNode` | ✅ |
| PDF10 | Metadata: `/Info` dict (`Title`/`Author` from `document.metadata` or `convert()` option overrides, omitted when absent — nothing fabricated; `Producer` always present); `pageSize`/`orientation`/`margins` → `MediaBox` via a points-native `resolvePageGeometry` (same option shape as DOCX's) | PDF1 | `Metadata` | ✅ |

**Acceptance:** All PDF1–10 milestones plus a whole-document acceptance
fixture (headings, mixed inline formatting, nested lists, a table, an image,
a hyperlink, enough filler to force multiple pages) pass structural
assertions against the real object graph and content-stream operators in
`tests/exporter-pdf.test.js`. Manually generated and inspected:
correct `%PDF-1.4` header/xref/trailer, correct page/font/image objects,
correctly separated bold/italic/code runs, WinAnsi-correct bullet and
smart-quote/dash bytes, a resolvable `/Link` annotation with the right
`/Rect`/`/URI`, and `/Info` metadata reflecting the given `title`/`author`.
Visual confirmation in a real PDF viewer (Adobe Reader / browser) is still
recommended but wasn't possible in this environment.

### Post-PDF10 — Box styling for paragraph/heading nodes ✅

PDF1–10 covered text content but none of the CSS "box styling" DOCX/PPTX
already read off `<div>`-collapsed paragraph/heading nodes (see those
exporters' own `shared/shape-style.js`, `resolveEdgeBorders`/`renderPBdr`,
`renderCardShapeDrawing`) — a real HTML report using `.hero`/`.card`-style
wrapper divs rendered as plain unstyled text with no background/border, and
(worse) a light-on-dark heading like `.hero{background:#1f4e79;color:#fff}`
rendered literally invisible (white text on the page's own white
background, no box drawn behind it at all).

`blocks.js`'s `renderStyledRunBlock`/`renderCardGroup`/`renderBlockSequence`
now read the same IDM fields DOCX/PPTX do and draw them as real PDF
geometry instead of OOXML/DrawingML markup:

- `backgroundColor` → a filled rect (or rounded-rect/ellipse, see below) behind the text
- `border` (uniform) → a stroked rect; per-side (`borderTop`/`Right`/`Bottom`/`Left`, e.g. a `border-left` accent stripe) → one thin filled rect per declared side
- `shadow` → an offset gray rect drawn first, behind the background/border (presence-only, like DOCX/PPTX's own fixed drop shadow — CSS's real offset/blur/color isn't representable)
- `rounded`/`circular` → a Bézier-curve rounded-rect or ellipse path (`content-stream.js`'s `roundedRect`/`ellipse`, PDF has no native rounded-rect operator, unlike DrawingML's `roundRect`/`ellipse` presets)
- `padding` → insets the wrapped text from the box edges on all 4 sides independently
- `margin.{beforePt,afterPt}` → replaces the flat inter-block gap when present
- `cardGroup` → several sibling paragraphs/headings that collapsed from one shared `<div>` container draw ONE box behind all of them, not one each
- `shrinkToFit` (`display:inline-block` or an explicit width) → the drawn box narrows to its own content width instead of stretching across the page

Also included: real CSS `display:grid`/`flex` column layout. `blocks.js`'s
`renderLayoutGroup`/`renderLayoutRow` read `layoutGroup`/`layoutItemIndex`
(dom-to-idm.js's `mapLayoutBlocks`) and lay sibling cells out genuinely
side by side — one `PdfPageFlow.withHorizontalSlot` per column, all starting
at the same row-top Y, with the shared cursor advancing by whichever cell
ended up tallest — instead of the flat vertical stack every node fell back
to before. Cells chunk into rows of `grid-template-columns`'s declared
count (or one single row for a plain `display:flex` with no declared
track count, the same "no wrapping, no unequal track sizes" simplification
PPTX's own `renderLayoutBlock`/DOCX's `renderLayoutTable` already accept). A
cell's required row height is pre-measured (`measureBlockSequenceHeight`,
sharing its math with the real renderer via `measureRunBlock`/
`measureCardGroup` rather than a separate approximation) so the whole row
can page-break cleanly instead of corrupting mid-row.

A grid cell whose source `<div>` produced more than one node (e.g. an
`<h3>`+bare-text pair, a very common "stat card" pattern) draws as ONE
shared box, not one per node — `renderLayoutGroup` synthesizes a cardGroup
across a cell's own members when at least one has box styling and none
already belongs to one. This matters because dom-to-idm's own cross-sibling
cardGroup heuristic requires a real `border`, or `backgroundColor` paired
with an explicit `color` (its `trustedPanel` gate) — a plain
`.box{background:#eef4ff}` with neither doesn't qualify, even though every
node it bled onto shares the identical background. Without this, each grid
cell's heading and value rendered as two visually disconnected boxes
stacked on top of each other instead of one card (each node independently
box-styled via renderStyledRunBlock's own single-node path).

Not attempted in this pass (a separate, larger follow-up): background
gradients as an actual multi-stop fill (currently only the existing
`backgroundColor` first-stop-color approximation), `anchorId` bookmarks,
`direction`/`writingMode`/`rotateDeg`, table cell `verticalAlign`, and box
styling on table/list/quote/code/section nodes themselves (only
paragraph/heading nodes — where this styling realistically lands from a
`<div>` collapse — are covered).

**Acceptance:** `tests/exporter-pdf.test.js`'s "box styling" describe block
(8 tests) verifies: a background renders as a filled rect drawn before its
text; a border-only box strokes without filling; `border-radius` produces
real Bézier `c` curve ops instead of a plain rect; padding measurably shifts
text position; `box-shadow` adds exactly one extra offset rect; a
multi-paragraph card shares exactly one drawn box; `display:inline-block`
narrows the box below full page width; and `margin-top`/`margin-bottom`
widen inter-block spacing beyond the default. Its "CSS grid/flex layout"
describe block (4 tests) verifies a 3-column grid's cells share one row at
increasing X positions, a plain `display:flex` row does the same with no
declared column count, a 2-column grid with 3 items wraps its 3rd cell onto
a second row, and a cell whose heading+value both inherit one wrapper's
background draws exactly one shared box (not one per node). A header
cell's own explicit background (PDF8) is also covered, with a sibling
unstyled header cell confirming the gray fallback still applies where no
explicit background exists. Manually verified against a representative
"annual report" fixture (hero banner, card, 3-box grid row, a styled table
header) that the previously-reported bugs are fixed: hero/card/box
backgrounds draw with their real CSS colors, white-on-colored-background
text is no longer invisible, the 3-box grid row renders as 3 side-by-side
columns instead of 6 stacked full-width bars, each grid cell's heading and
value share one box instead of two disconnected ones, and a table header
row uses its own declared background color instead of a flat gray default.

### Post-PDF10 — SVG vector shapes (DOCX + PPTX + PDF) ✅

An `<svg>` previously only ever mapped to content when it was purely a
wrapper around one embedded raster `<image>` (a common design-tool export
pattern) — genuine vector primitives (`<rect>`, `<circle>`, etc.) had no
representation at all and were silently dropped in every exporter. Fixed
for the "basic shapes" subset — `rect`, `circle`, `ellipse`, `line`,
`polygon`, `polyline` with `fill`/`stroke`/`stroke-width` — across all
three formats that have native vector-drawing primitives to target.

`dom-to-idm.js`'s `extractSvgShapes` (shared by every exporter) walks an
`<svg>`'s direct children — plus one level into any transform-less `<g>`,
pure grouping with no coordinate shift, which costs nothing to support —
extracting each shape's own geometry and resolved fill/stroke color (`style`
wins over a same-named bare attribute, matching real SVG precedence; `fill`
correctly defaults to black when undeclared, `stroke` to none — easy to get
backwards since every other property here defaults to "off"). Any declared
`viewBox` is resolved into the final display px box once, at parse time, so
no exporter needs to think about viewBox scaling at all. Produces a new
`svgShapes` IDM node (`idm/index.js`) carrying pre-scaled shape geometry.

Per-exporter rendering:
- **PDF** (`pdf/blocks.js`, `pdf/page-flow.js`): each shape draws via the
  same content-stream primitives already built for box-styling rounded
  corners (`drawRect`'s rounded/circular path, now reused for circle/
  ellipse) plus a new `drawPolygon` for polygon/polyline. SVG's Y-down
  coordinate space is flipped into PDF's Y-up page space.
- **PPTX** (`pptx/shapes.js`): each shape becomes its own absolutely
  positioned `<p:sp>` on the slide's own shape tree — PPTX already
  positions every shape via explicit EMU `a:off`/`a:ext`, the same
  convention SVG's own Y-down coordinates already use, so no coordinate
  flip or special compositing is needed. Polygon/polyline use a real
  `a:custGeom` custom path (this project's first use of DrawingML custom
  geometry); rect/circle/ellipse/line use `a:prstGeom` presets, reusing
  `shared/shape-style.js`'s existing `shapeFillXml`/`shapeLineXml`.
- **DOCX** (`document.js`): the one genuinely new mechanism this required —
  compositing several shapes at correct *relative* positions (e.g. a circle
  centered inside a rect) isn't possible with sequential `wp:inline`
  drawings (they just stack top-to-bottom). Each shape instead draws as its
  own floating `wp:anchor` — real ECMA-376 (not the `wpg:wgp` group-shape
  extension, deliberately avoided since it can't be verified against a real
  Word instance in this environment) — sharing one host paragraph, all
  positioned `relativeFrom="paragraph"`/`"column"` so their offsets compose
  correctly. The host paragraph's `w:spacing w:after` is set to the SVG's
  own height, since a floating drawing doesn't reserve layout space the way
  an inline one does — without it, the next paragraph would overlap the
  shapes instead of flowing below them. `resolveSvgShapeGeometry` (new,
  `shared/shape-style.js`) computes each shape's DrawingML geometry once,
  shared with PPTX's own (independently, already-shipped) shape math.

Not attempted (a separate, larger follow-up): `<path>` (arbitrary Bézier/
arc data), gradients on shapes, `<text>`/`<use>`/`<defs>`, filters/masks,
and any transformed `<g>` — each a substantially larger undertaking of its
own, left unguessed-at rather than approximated.

**Acceptance:** parser-level tests (`tests/html-parser.test.js`) verify
every shape kind's geometry extraction, fill/stroke defaulting and
`style`-wins-over-attribute precedence, viewBox rescaling (including a
non-zero origin), transform-less-`<g>` walking vs. transformed-`<g>`
skipping, and degenerate-shape dropping. Exporter-level tests confirm, for
each of PDF/PPTX/DOCX independently: a rect+circle logo renders as two
correctly-composited shapes (verified via center-point math, not just
presence) with the right fill colors; a viewBox rescale produces the right
final geometry; a polygon renders as a real closed custom-path shape with
both fill and stroke; a line renders stroke-only; and unsupported content
(e.g. a bare `<path>`) still drops silently with no crash. Manually
verified end-to-end against the reported rect+circle logo case in all
three formats with zero warnings.

### Post-PDF10 — Adjacent-margin collapsing ✅

Every pair of consecutive blocks that both declared vertical CSS margin
(a heading's `margin-bottom` flowing into a paragraph's `margin-top`, a
styled box's own `margin-top`/`margin-bottom`) rendered with roughly double
the intended gap — `renderStyledRunBlock`/`renderCardGroup` applied both
sides' margins as separate, unconditional `flow.advance()` calls, when real
CSS instead *collapses* two adjoining block-level margins into the larger
of the two, not their sum. Confirmed directly: a heading with
`margin-bottom:16pt` followed by a paragraph with `margin-top:16pt`
measurably rendered ~12pt further apart than the exact same heading
followed by a paragraph with `margin-top:0` — real collapsing makes those
identical (`max(16,0) === max(16,16)`), summing doesn't.

Fixed with a "pending spacing" model on `PdfPageFlow`
(`applySpacingBefore`/`deferSpacingAfter`/`resetPendingSpacing`): a block
*defers* its own trailing space instead of applying it immediately: the
*next* block's own leading space resolves the real gap as
`max(pending, ownBeforePt)`, applied exactly once. Wired into every
block-level renderer (heading/paragraph, cardGroup, quote, code, list,
table, thematic break, image, svgShapes) so a block with no real margin
concept (list/table/quote/...) still collapses its own flat default gap
against a real margin on either side, rather than adding an extra tax on
top of it. Grid/flex columns get special handling
(`resetPendingSpacing` before/after each cell in `renderLayoutRow`) since
they're independent side-by-side flows that must not leak deferred spacing
into each other despite sharing the same `PdfPageFlow` instance.

**Acceptance:** a dedicated test verifies a one-sided margin
(`margin-bottom:16pt` only) and a two-sided one
(`margin-bottom:16pt` + `margin-top:16pt`) produce the identical gap, and
that two *unequal* adjoining margins resolve to the larger one, not the
sum. Manually verified against the reported "too much vertical spacing
throughout the document" case.

---

## Suggested Build Order

```
Phase 0 (new primitives)
  P0.5 html-tokenizer  ← build first, gates Phase 1
  P0.6 zip-writer      ← build second, gates Phase 3
  P0.7 xml-builder     ← build third, gates Phase 3
        │
        ▼
Phase 1 — HTML Parser (H1 → H2 → H3 → H4 → H5 → H6)
        │
        ▼
Phase 2 — Text Exporters
  2A Markdown  ← simplest, proves pipeline end-to-end
  2B CSV
  2C TXT
  2D JSON
        │
        ▼
Phase 3 — OOXML Exporters
  3A DOCX  ← most complex OOXML; lessons transfer to XLSX/PPTX
  3B XLSX
  3C PPTX
        │
        ▼
Phase 4 — PDF (PDF1 → PDF2 → PDF3 → ... → PDF10, ship tier by tier)
```

---

## Definition of Done — per exporter milestone

A milestone is complete when:

1. It registers via the `Exporter` interface; selected by `format` id.
2. It reads only from IDM nodes — never directly from raw HTML.
3. Recoverable issues call `ctx.warn(code, message)`; fatal issues throw.
4. At least one golden snapshot test: known HTML fixture → verified output.
5. Output file opens correctly in its target application (Word, Excel, PDF viewer).
6. `npm run typecheck`, `npm test`, and `npm run lint` all pass.
7. For OOXML exporters (DOCX/XLSX/PPTX), at least one test parses the
   generated XML parts (via the project's own `zip.js`/`xml.js` readers) and
   checks structural conformance — namespace declarations present, grid/column
   consistency for tables, every `r:id`/`r:embed` reference resolving in the
   corresponding `.rels` file. A test that only checks the ZIP header or
   mimeType does not satisfy this — that's precisely what let a prior DOCX
   attempt reach "76 passing tests" while opening as corrupted in real Word.
