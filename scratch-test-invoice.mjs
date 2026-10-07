import { convert } from './src/index.js';
import { writeFile } from 'node:fs/promises';

const html = `
<table style="width:100%;">
<tr>
<td style="width:50%; vertical-align:top; padding:12px;">
<h3>Payment Details</h3>
<ul>
<li>Bank : HDFC Bank</li>
<li>Account : 123456789</li>
<li>IFSC : HDFC0000123</li>
<li>UPI : abc@upi</li>
</ul>
</td>
<td style="width:50%; vertical-align:top; padding:12px;">
<h3>Notes</h3>
<p>Payment is due within 7 days.</p>
<p>Late payments may incur additional charges.</p>
<p>Thank you for your business.</p>
</td>
</tr>
</table>
<h2 style="background:#1F4E79; color:white;">Terms & Conditions</h2>
<ol>
<li>Goods once sold will not be returned.</li>
</ol>
`;

const result = await convert(html, { format: 'pptx' });
await writeFile('./scratch-invoice.pptx', result.buffer);
console.log('done', result.buffer.length);
