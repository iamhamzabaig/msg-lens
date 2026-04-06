# msg-parser

## Project Goal

A universal npm package that parses .msg (Outlook) files and outputs sanitized HTML.
Must work in Node.js, React, Angular, Vue, and plain browser JS.

## Tech Stack

- TypeScript (strict mode)
- tsup for bundling (ESM + CJS dual output)
- Vitest for testing
- cfb for OLE2/CFBF binary parsing
- dompurify for HTML sanitization

## Folder Structure

src/
core/ # OLE2 binary reader, CFBF parser
mapi/ # MAPI property decoder
sanitizer/ # HTML sanitizer + cid: image resolver
types/ # All exported TypeScript interfaces
index.ts # Public API only
tests/
fixtures/ # .msg sample files for testing
unit/
integration/

## Coding Rules

- Never use Node.js-only APIs (Buffer, fs, path) in src/core or src/mapi
- Use Uint8Array and ArrayBuffer only — required for browser compatibility
- Every public function must have JSDoc comments
- Write a test for every new function before moving to the next
- Run `vitest` after each implementation step
- Run `tsc --noEmit` before committing anything

## Public API Contract

- Single entry: parseMsgFile(buffer: ArrayBuffer | Uint8Array): ParsedMessage
- Never throw uncaught errors — always return a typed result or typed error
- All HTML output must be XSS-safe (scripts, on\* attrs, external src stripped)

## Security Rules

- Strip <script> tags from all HTML bodies
- Strip all on\* event attributes
- Block external src/href (tracking pixels, remote resources)
- Replace cid: image references with base64 data URIs
- Handle malformed/corrupt .msg files gracefully — never crash

## Build Output

- ESM: dist/index.mjs
- CJS: dist/index.cjs
- Types: dist/index.d.ts
- Target: ES2020, no polyfills

## What NOT to Do

- Do not add runtime dependencies beyond cfb and dompurify
- Do not use any: use unknown with type guards instead
- Do not mix parsing logic into the public API layer (index.ts)
- Do not generate framework-specific code (no React hooks, no Angular services)
