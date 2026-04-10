# msg-lens

[![npm version](https://img.shields.io/npm/v/msg-lens?logo=npm)](https://www.npmjs.com/package/msg-lens)
[![npm downloads](https://img.shields.io/npm/dm/msg-lens?logo=npm)](https://www.npmjs.com/package/msg-lens)
[![MIT License](https://img.shields.io/badge/license-MIT-green.svg)](./LICENSE)
[![TypeScript](https://img.shields.io/badge/types-TypeScript-3178C6.svg)](https://www.typescriptlang.org/)

Universal Outlook `.msg` parser for Node.js and browsers.  
Parse once, get typed metadata + sanitized HTML ready to render.

## Package Links

- npm package: https://www.npmjs.com/package/msg-lens
- npm versions: https://www.npmjs.com/package/msg-lens?activeTab=versions
- npm files (tarball): https://www.npmjs.com/package/msg-lens?activeTab=code
- GitHub repository: https://github.com/iamhamzabaig/msg-lens
- Issues: https://github.com/iamhamzabaig/msg-lens/issues

## Highlights

- Cross-platform parsing: Node.js + browser support
- Single entrypoint: `parseMsgFile(data)`
- Typed result contract (`success` + structured payload/error)
- HTML sanitization for safe rendering
- Inline image resolution (`cid:` -> `data:` URI)
- Outlook RTF HTML extraction support
- Embedded `.msg` attachments support

## Installation

```bash
npm install msg-lens
```

## Quick Start

### Node.js

```ts
import { readFileSync } from 'fs';
import { parseMsgFile } from 'msg-lens';

const raw = readFileSync('email.msg');
const result = parseMsgFile(raw);

if (!result.success) {
  console.error(result.error.code, result.error.message);
  process.exit(1);
}

console.log('Subject:', result.message.subject);
console.log('From:', `${result.message.senderName} <${result.message.senderEmail}>`);
console.log('Attachments:', result.message.attachments.length);
```

### Browser

```ts
import { parseMsgFile } from 'msg-lens';

async function openMsg(file: File) {
  const buffer = await file.arrayBuffer();
  const result = parseMsgFile(buffer);

  if (!result.success) {
    console.error(result.error);
    return;
  }

  document.getElementById('preview')!.innerHTML = result.message.bodyHtml;
}
```

## API

### `parseMsgFile(data: ArrayBuffer | Uint8Array): ParseResult`

```ts
type ParseResult =
  | { success: true; message: ParsedMessage }
  | { success: false; error: ParseError };
```

### Core Types

`ParsedMessage` includes:

- `subject`, `senderName`, `senderEmail`
- `recipients`, `ccRecipients`, `bccRecipients`
- `bodyText`, `bodyHtml`
- `headers` (`date`, `dateObject`, `messageId`, `inReplyTo`, `importance`)
- `attachments` (`filename`, `mimeType`, `contentId`, `content`, `size`, `isInline`, optional `embeddedMessage`)

`ParseErrorCode`:

- `INVALID_CFB`
- `INVALID_EML`
- `MISSING_PROPERTIES`
- `MALFORMED_MAPI`
- `MALFORMED_MIME`
- `SANITIZATION_FAILED`
- `UNKNOWN_ERROR`

## Security

`bodyHtml` is sanitized before it is returned:

- strips `<script>` tags
- strips `on*` event handlers
- blocks `javascript:` URLs
- removes common 1x1 tracking pixels
- resolves `cid:` image references safely using attachment bytes

## Compatibility

- Input: `ArrayBuffer` and `Uint8Array`
- Output bundles:
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

## Release Workflow

```bash
npm run test
npm run typecheck
npm run build
npm version patch   # or minor / major
git push origin main
git push origin --tags
npm publish --access public
```

## License

MIT © msg-lens contributors. See [LICENSE](./LICENSE).
