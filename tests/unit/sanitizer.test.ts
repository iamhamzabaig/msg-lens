// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { sanitizeHtml } from '../../src/sanitizer';
import type { Attachment } from '../../src/types';

describe('sanitizeHtml', () => {
  it('returns empty string for empty input', () => {
    expect(sanitizeHtml('', [])).toBe('');
  });

  it('strips script tags', () => {
    const html = '<div>Hello</div><script>alert("xss")</script>';
    const result = sanitizeHtml(html, []);
    expect(result).not.toContain('<script');
    expect(result).toContain('Hello');
  });

  it('strips on* event attributes', () => {
    const html = '<img onerror="alert(1)" src="">';
    const result = sanitizeHtml(html, []);
    expect(result).not.toContain('onerror');
  });

  it('preserves style tags', () => {
    const html = '<style>.foo { color: red; }</style><div class="foo">Styled</div>';
    const result = sanitizeHtml(html, []);
    expect(result).toContain('Styled');
  });

  it('preserves inline style attributes', () => {
    const html = '<div style="color: red; font-size: 14px;">Styled text</div>';
    const result = sanitizeHtml(html, []);
    expect(result).toContain('style=');
    expect(result).toContain('Styled text');
  });

  it('preserves links with href', () => {
    const html = '<a href="https://example.com">Click here</a>';
    const result = sanitizeHtml(html, []);
    expect(result).toContain('href');
    expect(result).toContain('Click here');
  });

  it('preserves mailto: links', () => {
    const html = '<a href="mailto:test@example.com">Email</a>';
    const result = sanitizeHtml(html, []);
    expect(result).toContain('mailto:test@example.com');
  });

  it('strips 1x1 tracking pixels', () => {
    const html = '<img src="https://tracker.com/pixel.gif" width="1" height="1">';
    const result = sanitizeHtml(html, []);
    expect(result).not.toContain('tracker.com');
  });

  it('resolves cid: references to base64 data URIs', () => {
    const html = '<img src="cid:image001">';
    const attachment: Attachment = {
      filename: 'image.png',
      mimeType: 'image/png',
      contentId: 'image001',
      content: new Uint8Array([137, 80, 78, 71]),
      size: 4,
      isInline: true,
    };
    const result = sanitizeHtml(html, [attachment]);
    expect(result).toContain('data:image/png;base64,');
    expect(result).not.toContain('cid:');
  });

  it('resolves cid: references with angle brackets', () => {
    const html = '<img src="cid:img@domain.com">';
    const attachment: Attachment = {
      filename: 'photo.jpg',
      mimeType: 'image/jpeg',
      contentId: '<img@domain.com>',
      content: new Uint8Array([0xff, 0xd8]),
      size: 2,
      isInline: true,
    };
    const result = sanitizeHtml(html, [attachment]);
    expect(result).toContain('data:image/jpeg;base64,');
  });

  it('preserves table structure', () => {
    const html = '<table><tr><td style="padding: 10px;">Cell</td></tr></table>';
    const result = sanitizeHtml(html, []);
    expect(result).toContain('<table>');
    expect(result).toContain('<td');
  });

  it('strips javascript: URIs', () => {
    const html = '<a href="javascript:alert(1)">Click</a>';
    const result = sanitizeHtml(html, []);
    expect(result).not.toContain('javascript:');
  });
});
