/**
 * Quick manual test — pass a local file path OR a URL:
 *   node playground.mjs path/to/email.msg
 *   node playground.mjs https://example.com/email.msg
 */
import { readFileSync, writeFileSync } from 'fs';
import { parseMsgFile } from './dist/index.mjs';

const input = process.argv[2];
if (!input) {
  console.error('Usage: node playground.mjs <path-or-url-to-file.msg>');
  process.exit(1);
}

let buffer;
if (input.startsWith('http://') || input.startsWith('https://')) {
  console.log('Fetching from URL...');
  const res = await fetch(input);
  if (!res.ok) {
    console.error(`Fetch failed: ${res.status} ${res.statusText}`);
    process.exit(1);
  }
  buffer = new Uint8Array(await res.arrayBuffer());
} else {
  buffer = readFileSync(input);
}

const result = parseMsgFile(buffer);

if (!result.success) {
  console.error('Parse failed:', result.error);
  process.exit(1);
}

const msg = result.message;
console.log('Subject:', msg.subject);
console.log('From:', msg.senderName, `<${msg.senderEmail}>`);
console.log('To:', msg.recipients.map(r => `${r.name} <${r.email}>`).join(', '));
console.log('CC:', msg.ccRecipients.map(r => `${r.name} <${r.email}>`).join(', '));
console.log('Date:', msg.headers.date);
console.log('Attachments:', msg.attachments.length);
msg.attachments.forEach(a => console.log(`  - ${a.filename} (${a.mimeType}, ${a.size} bytes)`));

// Write HTML output so you can open it in a browser
writeFileSync('output.html', msg.bodyHtml);
console.log('\nHTML body written to output.html — open it in your browser.');
