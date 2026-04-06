import createDOMPurify from 'dompurify';
import type { Attachment } from '../types';

/**
 * Get a DOMPurify instance that works in both browser and Node.js.
 */
function getPurify(): { sanitize: (html: string, config: Record<string, unknown>) => string } | null {
  if (typeof window !== 'undefined') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return createDOMPurify(window as any);
  }
  if (typeof globalThis !== 'undefined' && (globalThis as Record<string, unknown>).window) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return createDOMPurify((globalThis as Record<string, unknown>).window as any);
  }
  return null;
}

/**
 * Sanitize HTML body and resolve cid: references to inline base64 data URIs.
 * @param html - Raw HTML body from the .msg file
 * @param attachments - Attachments array (used to resolve cid: references)
 * @returns Sanitized HTML string safe for rendering
 */
export function sanitizeHtml(html: string, attachments: Attachment[]): string {
  if (!html) return '';

  // Resolve cid: references to base64 data URIs
  let resolved = resolveCidReferences(html, attachments);

  // Try DOMPurify (available in browser and jsdom environments)
  const purify = getPurify();
  if (purify) {
    resolved = purify.sanitize(resolved, {
      ALLOWED_TAGS: [
        'html', 'head', 'body', 'div', 'span', 'p', 'br', 'hr',
        'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
        'a', 'img',
        'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'caption', 'colgroup', 'col',
        'ul', 'ol', 'li',
        'b', 'i', 'u', 'strong', 'em', 'small', 'sub', 'sup', 'mark',
        'blockquote', 'pre', 'code',
        'style', 'font', 'center',
      ],
      ALLOWED_ATTR: [
        'href', 'src', 'alt', 'title', 'width', 'height',
        'style', 'class', 'id', 'dir', 'lang',
        'colspan', 'rowspan', 'cellpadding', 'cellspacing', 'border',
        'align', 'valign', 'bgcolor', 'color', 'size', 'face',
        'target', 'rel',
      ],
      ALLOW_DATA_ATTR: false,
      FORBID_ATTR: ['onerror', 'onload', 'onclick', 'onmouseover'],
    });
  } else {
    // Fallback: regex-based sanitization for Node.js without jsdom
    resolved = regexSanitize(resolved);
  }

  // Only strip tracking pixel images, not links or styles
  resolved = stripTrackingPixels(resolved);

  return resolved;
}

/**
 * Regex-based fallback sanitizer for environments without a DOM.
 * Strips dangerous XSS vectors while preserving styles and layout.
 */
function regexSanitize(html: string): string {
  // Strip <script> tags and their content
  html = html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  // Strip on* event attributes (e.g. onerror, onclick)
  html = html.replace(/\s+on\w+\s*=\s*["'][^"']*["']/gi, '');
  html = html.replace(/\s+on\w+\s*=\s*[^\s>]+/gi, '');
  // Strip javascript: URIs in href/src
  html = html.replace(/(href|src)\s*=\s*["']\s*javascript:[^"']*["']/gi, '$1=""');
  // Strip <iframe>, <object>, <embed>, <form>, <input> tags
  html = html.replace(/<\/?(iframe|object|embed|form|input|textarea|button)\b[^>]*>/gi, '');
  return html;
}

/**
 * Replace cid:XXX references in HTML with base64 data URIs from attachments.
 */
function resolveCidReferences(html: string, attachments: Attachment[]): string {
  if (attachments.length === 0) return html;

  const cidMap = new Map<string, string>();
  for (const att of attachments) {
    if (att.contentId) {
      const cleanCid = att.contentId.replace(/^<|>$/g, '');
      const base64 = uint8ArrayToBase64(att.content);
      const dataUri = `data:${att.mimeType};base64,${base64}`;
      cidMap.set(cleanCid, dataUri);
    }
  }

  if (cidMap.size === 0) return html;

  return html.replace(/src\s*=\s*["']cid:([^"']+)["']/gi, (_match, cid) => {
    const dataUri = cidMap.get(cid);
    return dataUri ? `src="${dataUri}"` : _match;
  });
}

/**
 * Strip 1x1 tracking pixel images (external tiny images used for email tracking).
 * Preserves normal images and all other elements.
 */
function stripTrackingPixels(html: string): string {
  // Remove img tags that look like tracking pixels (1x1, hidden, or zero-size)
  html = html.replace(
    /<img\b[^>]*(?:width\s*=\s*["']?1["']?\s+height\s*=\s*["']?1["']?|height\s*=\s*["']?1["']?\s+width\s*=\s*["']?1["']?)[^>]*>/gi,
    '',
  );
  return html;
}

/**
 * Convert Uint8Array to base64 string (browser-safe, no Buffer).
 */
function uint8ArrayToBase64(bytes: Uint8Array): string {
  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}
