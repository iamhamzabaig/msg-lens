# msg-lens

Universal `.msg` (Outlook) email parser for Node.js and browsers.

`msg-lens` parses Outlook `.msg` files into a typed object and returns HTML that is sanitized and ready to render.

## Why msg-lens

- Works in Node.js and modern browsers
- Single API: `parseMsgFile(data)`
- Never throws uncaught parse errors (typed result wrapper)
- Sanitizes HTML output (XSS-safe defaults)
- Resolves inline `cid:` images to base64 data URIs
- Supports Outlook RTF-with-HTML payload extraction
- Handles embedded `.msg` attachments (recursive with depth limit)

## Install

```bash
npm install msg-lens
```

## Quick Start

### Node.js

```ts
import { readFileSync } from 'fs';
import { parseMsgFile } from 'msg-lens';

const data = readFileSync('email.msg');
const result = parseMsgFile(data);

if (!result.success) {
  console.error(result.error.code, result.error.message);
  process.exit(1);
}

console.log('Subject:', result.message.subject);
console.log('From:', `${result.message.senderName} <${result.message.senderEmail}>`);
console.log('Attachments:', result.message.attachments.length);
```

### Browser (file input)

```ts
import { parseMsgFile } from 'msg-lens';

const input = document.querySelector<HTMLInputElement>('#file');
const preview = document.querySelector<HTMLIFrameElement>('#preview');

input?.addEventListener('change', () => {
  const file = input.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = () => {
    const result = parseMsgFile(reader.result as ArrayBuffer);
    if (!result.success) {
      console.error(result.error);
      return;
    }
    if (preview) preview.srcdoc = result.message.bodyHtml;
  };
  reader.readAsArrayBuffer(file);
});
```

## API

### `parseMsgFile(data: ArrayBuffer | Uint8Array): ParseResult`

```ts
type ParseResult =
  | { success: true; message: ParsedMessage }
  | { success: false; error: ParseError };
```

### `ParsedMessage`

| Field | Type | Notes |
| --- | --- | --- |
| `subject` | `string` | Message subject |
| `senderName` | `string` | Sender display name |
| `senderEmail` | `string` | Sender email |
| `recipients` | `Recipient[]` | To recipients |
| `ccRecipients` | `Recipient[]` | CC recipients |
| `bccRecipients` | `Recipient[]` | BCC recipients |
| `bodyText` | `string` | Plain text body |
| `bodyHtml` | `string` | Sanitized HTML body |
| `headers` | `MessageHeaders` | Date, message ID, importance |
| `attachments` | `Attachment[]` | Attachments and inline resources |

### `Recipient`

| Field | Type |
| --- | --- |
| `name` | `string` |
| `email` | `string` |
| `type` | `'to' \| 'cc' \| 'bcc'` |

### `MessageHeaders`

| Field | Type |
| --- | --- |
| `date` | `string` |
| `dateObject` | `Date \| null` |
| `messageId` | `string` |
| `inReplyTo` | `string` |
| `importance` | `'low' \| 'normal' \| 'high'` |

### `Attachment`

| Field | Type | Notes |
| --- | --- | --- |
| `filename` | `string` | Attachment filename |
| `mimeType` | `string` | MIME type |
| `contentId` | `string` | Used for inline `cid:` references |
| `content` | `Uint8Array` | Raw bytes |
| `size` | `number` | Byte length |
| `isInline` | `boolean` | Inline/embedded indicator |
| `embeddedMessage` | `ParsedMessage \| undefined` | Present for embedded `.msg` attachments |

### `ParseError`

| Field | Type |
| --- | --- |
| `code` | `ParseErrorCode` |
| `message` | `string` |

`ParseErrorCode` values:

- `INVALID_CFB`
- `INVALID_EML`
- `MISSING_PROPERTIES`
- `MALFORMED_MAPI`
- `MALFORMED_MIME`
- `SANITIZATION_FAILED`
- `UNKNOWN_ERROR`

## Security Model

`bodyHtml` is sanitized before returning from `parseMsgFile`.

Current protections include:

- removes `<script>` tags
- strips inline event handlers (`on*`)
- blocks `javascript:` URLs
- strips known tracking-pixel patterns (1x1 images)
- rewrites `cid:` image references to safe `data:` URIs using attachment bytes

## Runtime Compatibility

- Browser-safe parser code (`ArrayBuffer` / `Uint8Array`)
- Dual package output:
  - ESM: `dist/index.mjs`
  - CJS: `dist/index.js`
  - Types: `dist/index.d.ts`

## Development

```bash
npm install
npm run test
npm run typecheck
npm run build
```

Useful local files:

- `sandbox.html` for browser preview/testing
- `playground.mjs` for quick Node parsing tests

## Release Checklist (Git + npm)

### 1) Validate locally

```bash
npm run test
npm run typecheck
npm run build
```

### 2) Update version

Choose one:

```bash
npm version patch
npm version minor
npm version major
```

This updates `package.json`, `package-lock.json`, creates a commit, and tags the release.

### 3) Push code and tags

```bash
git push origin main
git push origin --tags
```

### 4) Publish to npm

```bash
npm publish --access public
```

### 5) Verify published version

```bash
npm view msg-lens version
```

## License

ISC
