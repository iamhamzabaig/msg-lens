import { parseCfb } from './core';
import { decodeMessage } from './mapi';
import { sanitizeHtml } from './sanitizer';
import type { ParseResult, ParsedMessage } from './types';

// Re-export types for consumers
export type {
  ParsedMessage,
  Recipient,
  Attachment,
  MessageHeaders,
  ParseResult,
  ParseError,
  ParseErrorCode,
} from './types';

/**
 * Parse a .msg (Outlook) file buffer and return a structured, sanitized result.
 *
 * @param data - Raw .msg file as ArrayBuffer or Uint8Array
 * @returns A typed result: `{ success: true, message }` or `{ success: false, error }`
 *
 * @example
 * ```ts
 * const response = await fetch('https://example.com/email.msg');
 * const buffer = await response.arrayBuffer();
 * const result = parseMsgFile(buffer);
 * if (result.success) {
 *   document.innerHTML = result.message.bodyHtml;
 * }
 * ```
 */
export function parseMsgFile(data: ArrayBuffer | Uint8Array): ParseResult {
  try {
    // Step 1: Parse CFB container
    const streams = parseCfb(data);
    if (!streams) {
      return {
        success: false,
        error: { code: 'INVALID_CFB', message: 'Failed to parse CFB container. The file may be corrupt or not a valid .msg file.' },
      };
    }

    if (streams.length === 0) {
      return {
        success: false,
        error: { code: 'MISSING_PROPERTIES', message: 'No streams found in CFB container.' },
      };
    }

    // Step 2: Decode MAPI properties
    let message: ParsedMessage;
    try {
      message = decodeMessage(streams);
    } catch {
      return {
        success: false,
        error: { code: 'MALFORMED_MAPI', message: 'Failed to decode MAPI properties from .msg file.' },
      };
    }

    // Step 3: Sanitize HTML body and resolve cid: images
    try {
      if (message.bodyHtml) {
        message = {
          ...message,
          bodyHtml: sanitizeHtml(message.bodyHtml, message.attachments),
        };
      } else if (message.bodyText) {
        // Convert plain text to basic HTML
        const escaped = message.bodyText
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/\n/g, '<br>');
        message = {
          ...message,
          bodyHtml: `<div>${escaped}</div>`,
        };
      }
    } catch {
      return {
        success: false,
        error: { code: 'SANITIZATION_FAILED', message: 'Failed to sanitize HTML body.' },
      };
    }

    return { success: true, message };
  } catch {
    return {
      success: false,
      error: { code: 'UNKNOWN_ERROR', message: 'An unexpected error occurred while parsing the .msg file.' },
    };
  }
}
