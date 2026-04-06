# msg-lens

A universal `.msg` (Outlook) file parser that extracts email content and outputs sanitized HTML. Works in **Node.js**, **React**, **Angular**, **Vue**, and **plain browser JS**.

## Install

```bash
npm install msg-lens
```

## Quick Start

```ts
import { parseMsgFile } from 'msg-lens';

// From a URL
const response = await fetch('https://example.com/email.msg');
const buffer = await response.arrayBuffer();
const result = parseMsgFile(buffer);

if (result.success) {
  console.log(result.message.subject);
  console.log(result.message.senderName);
  console.log(result.message.bodyHtml); // sanitized HTML ready to render
}
```

```ts
// From a file input (browser)
const input = document.querySelector('input[type="file"]');
input.addEventListener('change', (e) => {
  const reader = new FileReader();
  reader.onload = () => {
    const result = parseMsgFile(reader.result);
    if (result.success) {
      document.getElementById('preview').innerHTML = result.message.bodyHtml;
    }
  };
  reader.readAsArrayBuffer(e.target.files[0]);
});
```

## API

### `parseMsgFile(data: ArrayBuffer | Uint8Array): ParseResult`

Parses a `.msg` file and returns a typed result.

```ts
type ParseResult =
  | { success: true; message: ParsedMessage }
  | { success: false; error: ParseError };
```

### ParsedMessage

| Field | Type | Description |
|-------|------|-------------|
| `subject` | `string` | Email subject line |
| `senderName` | `string` | Sender display name |
| `senderEmail` | `string` | Sender email address |
| `recipients` | `Recipient[]` | To recipients |
| `ccRecipients` | `Recipient[]` | CC recipients |
| `bccRecipients` | `Recipient[]` | BCC recipients |
| `bodyText` | `string` | Plain text body |
| `bodyHtml` | `string` | Sanitized HTML body (XSS-safe) |
| `headers` | `MessageHeaders` | Date, message ID, importance |
| `attachments` | `Attachment[]` | File attachments |

### Recipient

| Field | Type |
|-------|------|
| `name` | `string` |
| `email` | `string` |
| `type` | `'to' \| 'cc' \| 'bcc'` |

### Attachment

| Field | Type | Description |
|-------|------|-------------|
| `filename` | `string` | Original filename |
| `mimeType` | `string` | MIME type |
| `contentId` | `string` | Content-ID for inline images |
| `content` | `Uint8Array` | Raw file bytes |
| `size` | `number` | Size in bytes |
| `isInline` | `boolean` | Whether it's an embedded image |

### ParseError

| Field | Type |
|-------|------|
| `code` | `'INVALID_CFB' \| 'MISSING_PROPERTIES' \| 'MALFORMED_MAPI' \| 'SANITIZATION_FAILED' \| 'UNKNOWN_ERROR'` |
| `message` | `string` |

## Framework Examples

### React

```tsx
import { parseMsgFile, ParsedMessage } from 'msg-lens';
import { useState } from 'react';

function MsgViewer() {
  const [email, setEmail] = useState<ParsedMessage | null>(null);

  const handleFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const result = parseMsgFile(reader.result as ArrayBuffer);
      if (result.success) setEmail(result.message);
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div>
      <input type="file" accept=".msg" onChange={handleFile} />
      {email && (
        <>
          <h2>{email.subject}</h2>
          <p>From: {email.senderName}</p>
          <iframe srcDoc={email.bodyHtml} style={{ width: '100%', height: 600 }} />
        </>
      )}
    </div>
  );
}
```

### Vue

```vue
<script setup lang="ts">
import { ref } from 'vue';
import { parseMsgFile, ParsedMessage } from 'msg-lens';

const email = ref<ParsedMessage | null>(null);

function handleFile(e: Event) {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const result = parseMsgFile(reader.result as ArrayBuffer);
    if (result.success) email.value = result.message;
  };
  reader.readAsArrayBuffer(file);
}
</script>

<template>
  <input type="file" accept=".msg" @change="handleFile" />
  <div v-if="email">
    <h2>{{ email.subject }}</h2>
    <p>From: {{ email.senderName }}</p>
    <iframe :srcdoc="email.bodyHtml" style="width: 100%; height: 600px" />
  </div>
</template>
```

### Node.js

```ts
import { readFileSync } from 'fs';
import { parseMsgFile } from 'msg-lens';

const buffer = readFileSync('email.msg');
const result = parseMsgFile(buffer);

if (result.success) {
  console.log('Subject:', result.message.subject);
  console.log('From:', result.message.senderName, result.message.senderEmail);
  console.log('Attachments:', result.message.attachments.length);
}
```

## Features

- **Universal** — works in Node.js and all browsers (no `Buffer` or `fs` dependency)
- **Single function API** — `parseMsgFile()` does everything
- **XSS-safe HTML** — scripts, event handlers, and tracking pixels stripped
- **Inline images** — `cid:` references resolved to base64 data URIs
- **RTF support** — extracts HTML from Outlook's compressed RTF format
- **Smart quotes** — Windows-1252 characters decoded correctly
- **Never crashes** — returns typed errors instead of throwing
- **Tiny** — only 2 runtime dependencies (`cfb` + `dompurify`)
- **TypeScript** — full type definitions included
- **Dual format** — ships ESM + CJS

## Security

All HTML output is sanitized:

- `<script>` tags removed
- `on*` event attributes stripped
- `javascript:` URIs blocked
- 1x1 tracking pixel images removed
- `cid:` image references replaced with safe base64 data URIs

## License

ISC
