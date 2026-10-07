// Run with: node examples/basic.mjs
import { convert } from '../src/index.js';
import { writeFile } from 'node:fs/promises';

const html = `
<html lang="en">
<head><title>Quarterly Report</title></head>
<body>
<h1>Quarterly Report</h1>
<p>A short <strong>summary</strong> with a <a href="https://example.com">link</a>.</p>

<ul>
  <li>First point</li>
  <li>Second point</li>
</ul>

<blockquote><p>A closing remark.</p></blockquote>

<table>
  <thead><tr><th>Item</th><th>Value</th></tr></thead>
  <tbody>
    <tr><td>Revenue</td><td>$1M</td></tr>
    <tr><td>Costs</td><td>$600K</td></tr>
  </tbody>
</table>
</body>
</html>
`;

// Export to Markdown
const md = await convert(html, { format: 'markdown' });
console.log('--- MARKDOWN ---');
console.log(md.buffer.toString('utf8'));

// Export to plain text
const txt = await convert(html, { format: 'txt' });
console.log('--- PLAIN TEXT ---');
console.log(txt.buffer.toString('utf8'));

// Export to DOCX (binary — save to file)
const docx = await convert(html, { format: 'docx', title: 'Quarterly Report', author: 'Finance Team' });
await writeFile('output.docx', docx.buffer);
console.log(`DOCX saved → output.docx (${docx.buffer.length} bytes)`);

// Export to PDF (binary — save to file)
const pdf = await convert(html, { format: 'pdf' });
await writeFile('output.pdf', pdf.buffer);
console.log(`PDF  saved → output.pdf (${pdf.buffer.length} bytes)`);

// Export to JSON
const json = await convert(html, { format: 'json' });
console.log('--- JSON ---');
const data = JSON.parse(json.buffer.toString('utf8'));
console.log(`Title: ${data.meta.title}`);
console.log(`Blocks: ${data.content.length}`);
