# Agents Guide for `msg-lens`

This document is for coding agents working in this repository. It captures the real implementation state (not just intended behavior) so changes stay safe, compatible, and reviewable.

## 1) Repository Snapshot

- Package: `msg-lens` (`1.0.1`)
- Goal: Parse Outlook `.msg` files and output structured message data plus sanitized HTML.
- Runtime deps: `cfb`, `dompurify`
- Tooling: TypeScript (strict), tsup (ESM+CJS+d.ts), Vitest
- Verified on this repo state:
  - `npm run test` -> 22 tests passing
  - `npm run typecheck` -> passing
  - `npm run build` -> passing

## 2) High-Level Architecture

Pipeline in public API:

1. `parseMsgFile` in `src/index.ts`
2. `parseCfb` in `src/core/cfb-reader.ts`
3. `decodeMessage` in `src/mapi/decoder.ts`
4. `sanitizeHtml` in `src/sanitizer/index.ts`
5. Return typed `ParseResult`

Core principle: library code is browser-compatible and avoids Node-only APIs in `src/`.

## 3) Module Map and Responsibilities

- `src/index.ts`
  - Public entrypoint and error mapping.
  - Converts `bodyText` to basic escaped HTML if no HTML body exists.
- `src/core/cfb-reader.ts`
  - Wraps `cfb` parser and extracts stream entries only (`entry.type === 2`).
  - Provides `findStream` and `findStreams` helpers.
- `src/core/rtf-decompress.ts`
  - Implements LZFu decompression for compressed RTF.
  - Extracts Outlook `\fromhtml` content via `\htmltag` parsing.
- `src/mapi/decoder.ts`
  - Main MAPI property decoding.
  - Parses sender, recipients, headers, attachments.
  - Handles embedded `.msg` attachments recursively (`ATTACH_EMBEDDED_MSG`, depth-limited).
- `src/mapi/property-tags.ts`
  - MAPI constants and attachment/recipient enums.
- `src/sanitizer/index.ts`
  - Resolves `cid:` references to base64 data URIs.
  - Sanitizes HTML via DOMPurify if window exists, else regex fallback.
  - Removes 1x1 tracking pixel `<img>` tags.
- `src/types/index.ts`
  - Full exported domain model and parse error contract.

## 4) Data and Error Contract Invariants

Must preserve:

- `parseMsgFile` never throws; always returns `{ success: true|false, ... }`.
- Input compatibility: `ArrayBuffer | Uint8Array`.
- `ParsedMessage.bodyHtml` is sanitized output intended for direct rendering.
- Inline CID images are resolved from attachment `contentId`.
- Embedded message recursion must remain bounded (`MAX_NESTING_DEPTH`).

Error handling shape:

- Runtime currently returns: `INVALID_CFB`, `MISSING_PROPERTIES`, `MALFORMED_MAPI`, `SANITIZATION_FAILED`, `UNKNOWN_ERROR`.
- `ParseErrorCode` type also includes `INVALID_EML` and `MALFORMED_MIME` (currently unused by runtime path).

## 5) Known Gaps / Drift (Important)

These are current repo realities and should be considered before new work:

- API drift in docs vs code:
  - README error union omits `INVALID_EML` and `MALFORMED_MIME` present in `src/types/index.ts`.
  - `CLAUDE.md` says CJS output is `dist/index.cjs`, but build currently emits `dist/index.js`.
- Limited test coverage:
  - No tests for `src/mapi/decoder.ts` behavior on real `.msg` fixtures.
  - No tests for `src/core/rtf-decompress.ts`.
  - No tests for embedded message recursion and depth cap.
- Sanitizer behavior depends on environment:
  - DOMPurify path in browser/jsdom.
  - Regex fallback path when no `window` is available.

## 6) Change Playbooks

### A) If editing public API (`src/index.ts`)

- Keep return type stable (`ParseResult`).
- Preserve no-throw guarantee.
- Re-validate fallback behavior when `bodyHtml` is absent and `bodyText` exists.
- Update README examples if function signature/behavior changes.

### B) If editing MAPI decode logic (`src/mapi/decoder.ts`)

- Confirm property tag changes in `src/mapi/property-tags.ts`.
- Keep Unicode-first then STRING8 fallback for strings.
- Verify recipient classification (`to`/`cc`/`bcc`) still maps MAPI values correctly.
- For attachments:
  - Keep inline detection logic (`contentId`/rendering position).
  - Re-check embedded message recursion depth safeguards.

### C) If editing RTF logic (`src/core/rtf-decompress.ts`)

- Preserve LZFu header checks and dictionary initialization.
- Validate `\htmlrtf` skip-state and `\htmltag` extraction.
- Add regression tests for malformed or edge-case RTF input.

### D) If editing sanitizer (`src/sanitizer/index.ts`)

- Preserve security baseline:
  - no script tags
  - no inline event handlers
  - no `javascript:` URIs
- Ensure `cid:` resolution still works with angle-bracket content IDs.
- Be careful with false positives in tracking-pixel stripping.

### E) If editing types/docs (`src/types`, `README.md`, `CLAUDE.md`)

- Keep docs synchronized with actual emitted artifacts and runtime behavior.
- If adding parse error variants, update:
  - `ParseErrorCode` type
  - runtime error mapping in `parseMsgFile`
  - README API section

## 7) Test Strategy Expectations

Minimum checks before handing off:

1. `npm run test`
2. `npm run typecheck`
3. `npm run build`

When touching specific areas:

- `mapi/decoder.ts`: add/update unit tests with representative stream structures.
- `rtf-decompress.ts`: add focused decompression/extraction tests (including malformed data).
- `sanitizer/index.ts`: test both DOMPurify behavior and regex fallback behavior when possible.

## 8) Security and Compatibility Rules

- Maintain browser-safe implementation in `src/` (no `fs`, `path`, Node `Buffer` in library runtime).
- Keep HTML output sanitized before returning it from public API.
- Do not introduce heavy or unnecessary runtime dependencies.

## 9) Practical Working Notes for Agents

- Prefer small, isolated changes with tests in same PR/commit.
- Do not edit `dist/` manually; regenerate via build.
- Assume working tree may already be dirty; do not revert unrelated changes.
- If behavior changes intentionally, record it in README and tests together.

## 10) Suggested Near-Term Improvements

1. Add fixture-based integration tests for real `.msg` files.
2. Add unit coverage for RTF decompression and embedded `.msg` recursion.
3. Resolve docs/type/runtime drift in parse error codes and build output naming.
4. Add a compatibility matrix (Node versions + browser smoke check).
