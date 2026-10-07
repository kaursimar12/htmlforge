// A richer example than basic.mjs — styled boxes, a table with real CSS
// padding/border, a code block, a genuinely-styled note box vs. a plain
// blockquote (no invented styling), a nested list, an embedded image, and
// exporting the same HTML to five different formats with PDF-specific page
// options.
//
// Run with: node examples/report.mjs
import { convert } from '../src/index.js';
import { writeFile } from 'node:fs/promises';

// A tiny 8x8 navy-blue PNG, embedded as a base64 data: URI — no external
// file or network fetch needed to run this example.
const IMAGE_DATA_URI =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGOQ96vEihiGlgQAYZI5gapHhb8AAAAASUVORK5CYII=';

const html = `
<html lang="en">
<head>
  <title>Quarterly Report</title>
  <style>
    body { font: 15px/1.5 Arial, sans-serif; }
    .hero { background: #1f4e79; color: #fff; padding: 24px; border-radius: 8px; }
    .note { border-left: 4px solid #1f4e79; background: #eef6ff; padding: 10px; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid #999; padding: 6px; }
    th { background: #ddebf7; }
    pre { background: #272822; color: #eee; padding: 10px; }
  </style>
</head>
<body>
  <div class="hero">
    <h1>Quarterly Report</h1>
    <p>Prepared for stakeholders</p>
  </div>

  <h2>Summary</h2>
  <p>
    Revenue is <strong>up</strong>, costs are <em>down</em>, and the team shipped
    <mark>three</mark> major features. See H<sub>2</sub>O for a chemistry joke,
    and read more at <a href="https://example.com">example.com</a>.
    Code sample: <code>npm run build</code>.
  </p>

  <div class="note"><strong>Note:</strong> Figures are provisional pending audit.</div>

  <blockquote>A plain, unstyled quote — no background or border in the CSS above, so it renders as plain indented text, not an invented highlight box.</blockquote>

  <h2>Numbers</h2>
  <table>
    <tr><th>Metric</th><th>Q1</th><th>Q2</th></tr>
    <tr><td>Revenue</td><td>$1.2M</td><td>$1.6M</td></tr>
    <tr><td>Costs</td><td>$600K</td><td>$550K</td></tr>
  </table>

  <h2>Highlights</h2>
  <ul>
    <li>Improved onboarding flow</li>
    <li>Expanded platform support
      <ul>
        <li>iOS</li>
        <li>Android</li>
      </ul>
    </li>
  </ul>

  <h2>Architecture</h2>
  <pre>HTML → IDM → Exporter → Document</pre>

  <img src="${IMAGE_DATA_URI}" alt="logo" width="64" height="64">
</body>
</html>
`;

// PDF with explicit page options — pageSize/orientation/margins all in mm.
const pdf = await convert(html, {
  format: 'pdf',
  title: 'Quarterly Report',
  author: 'Finance Team',
  pageSize: 'A4',
  orientation: 'portrait',
  margins: { top: 20, right: 20, bottom: 20, left: 20 },
});
await writeFile('report.pdf', pdf.buffer);
console.log(`PDF  saved -> report.pdf (${pdf.buffer.length} bytes)`);
if (pdf.warnings.length) console.log('  warnings:', pdf.warnings.map((w) => w.code));

const docx = await convert(html, { format: 'docx', title: 'Quarterly Report', author: 'Finance Team' });
await writeFile('report.docx', docx.buffer);
console.log(`DOCX saved -> report.docx (${docx.buffer.length} bytes)`);

const pptx = await convert(html, { format: 'pptx', title: 'Quarterly Report' });
await writeFile('report.pptx', pptx.buffer);
console.log(`PPTX saved -> report.pptx (${pptx.buffer.length} bytes)`);

const xlsx = await convert(html, { format: 'xlsx', sheetName: 'Numbers' });
await writeFile('report.xlsx', xlsx.buffer);
console.log(`XLSX saved -> report.xlsx (${xlsx.buffer.length} bytes)`);

const md = await convert(html, { format: 'markdown' });
console.log('\n--- MARKDOWN ---');
console.log(md.buffer.toString('utf8'));
