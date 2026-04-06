import type { CfbStream } from '../core';
import { decompressRtf, extractHtmlFromRtf } from '../core';
import type {
  ParsedMessage,
  Recipient,
  Attachment,
  MessageHeaders,
} from '../types';
import {
  PT_UNICODE,
  PT_STRING8,
  PT_BINARY,
  PT_LONG,
  PT_BOOLEAN,
  PidTagSubject,
  PidTagSenderName,
  PidTagSenderEmailAddress,
  PidTagSenderSmtpAddress,
  PidTagBody,
  PidTagHtml,
  PidTagRtfCompressed,
  PidTagMessageDeliveryTime,
  PidTagInternetMessageId,
  PidTagInReplyToId,
  PidTagImportance,
  PidTagDisplayName,
  PidTagEmailAddress,
  PidTagSmtpAddress,
  PidTagRecipientType,
  PidTagAttachFilename,
  PidTagAttachLongFilename,
  PidTagAttachMimeTag,
  PidTagAttachContentId,
  PidTagAttachDataBinary,
  PidTagRenderingPosition,
  MAPI_TO,
  MAPI_CC,
  MAPI_BCC,
} from './property-tags';

/**
 * Decode a .msg file's CFB streams into a ParsedMessage.
 * @param streams - All streams extracted from the CFB container
 */
export function decodeMessage(streams: CfbStream[]): ParsedMessage {
  const subject = getStringProperty(streams, '', PidTagSubject);
  const senderName = getStringProperty(streams, '', PidTagSenderName);
  const senderEmail =
    getStringProperty(streams, '', PidTagSenderSmtpAddress) ||
    getStringProperty(streams, '', PidTagSenderEmailAddress);

  const bodyText = getStringProperty(streams, '', PidTagBody);
  const bodyHtmlRaw = getHtmlBody(streams);

  const headers = decodeHeaders(streams);
  const recipients = decodeRecipients(streams);
  const attachments = decodeAttachments(streams);

  return {
    subject,
    senderName,
    senderEmail,
    recipients: recipients.filter(r => r.type === 'to'),
    ccRecipients: recipients.filter(r => r.type === 'cc'),
    bccRecipients: recipients.filter(r => r.type === 'bcc'),
    bodyText,
    bodyHtml: bodyHtmlRaw,
    headers,
    attachments,
  };
}

// --- Property stream reading ---

/**
 * .msg files store properties in streams named like:
 *   __substg1.0_PPPPTTTTT
 * where PPPP = property ID (hex), TTTT = property type (hex).
 *
 * For the root message, the prefix is "Root Entry/" or just the root.
 * For recipients: "__recip_version1.0_#XXXXXXXX/"
 * For attachments: "__attach_version1.0_#XXXXXXXX/"
 */

function buildStreamName(propId: number, propType: number): string {
  const id = propId.toString(16).toUpperCase().padStart(4, '0');
  const type = propType.toString(16).toUpperCase().padStart(4, '0');
  return `__substg1.0_${id}${type}`;
}

function findPropertyStream(
  streams: CfbStream[],
  prefix: string,
  propId: number,
  propType: number,
): CfbStream | undefined {
  const streamName = buildStreamName(propId, propType);
  return streams.find(s => {
    const normalizedPath = s.path.replace(/\\/g, '/');
    if (prefix) {
      return normalizedPath.includes(prefix) && s.name === streamName;
    }
    // Root-level: stream directly under root entry
    const parts = normalizedPath.split('/');
    return parts.length <= 2 && s.name === streamName;
  });
}

function getStringProperty(
  streams: CfbStream[],
  prefix: string,
  propId: number,
): string {
  // Try Unicode first, then ASCII
  const unicodeStream = findPropertyStream(streams, prefix, propId, PT_UNICODE);
  if (unicodeStream && unicodeStream.content.length > 0) {
    return decodeUtf16Le(unicodeStream.content);
  }

  const asciiStream = findPropertyStream(streams, prefix, propId, PT_STRING8);
  if (asciiStream && asciiStream.content.length > 0) {
    return decodeAscii(asciiStream.content);
  }

  return '';
}

function getBinaryProperty(
  streams: CfbStream[],
  prefix: string,
  propId: number,
): Uint8Array | null {
  const stream = findPropertyStream(streams, prefix, propId, PT_BINARY);
  return stream ? stream.content : null;
}

function getLongProperty(
  streams: CfbStream[],
  prefix: string,
  propId: number,
): number | null {
  // Long properties may be stored in the properties stream or as separate streams
  const stream = findPropertyStream(streams, prefix, propId, PT_LONG);
  if (stream && stream.content.length >= 4) {
    return readUint32Le(stream.content, 0);
  }
  return null;
}

// --- String decoding (browser-safe, no Buffer) ---

function decodeUtf16Le(data: Uint8Array): string {
  const chars: string[] = [];
  for (let i = 0; i + 1 < data.length; i += 2) {
    const code = data[i] | (data[i + 1] << 8);
    if (code === 0) break;
    chars.push(String.fromCharCode(code));
  }
  return chars.join('');
}

function decodeAscii(data: Uint8Array): string {
  const chars: string[] = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i] === 0) break;
    chars.push(String.fromCharCode(data[i]));
  }
  return chars.join('');
}

function readUint32Le(data: Uint8Array, offset: number): number {
  return (
    data[offset] |
    (data[offset + 1] << 8) |
    (data[offset + 2] << 16) |
    ((data[offset + 3] << 24) >>> 0)
  );
}

// --- HTML body extraction ---

function getHtmlBody(streams: CfbStream[]): string {
  // Try HTML binary property first (PidTagHtml 0x1013)
  const htmlBinary = getBinaryProperty(streams, '', PidTagHtml);
  if (htmlBinary && htmlBinary.length > 0) {
    return decodeUtf8(htmlBinary);
  }

  // Try compressed RTF with embedded HTML (PidTagRtfCompressed 0x1009)
  const rtfBinary = getBinaryProperty(streams, '', PidTagRtfCompressed);
  if (rtfBinary && rtfBinary.length > 0) {
    const rtf = decompressRtf(rtfBinary);
    if (rtf) {
      const html = extractHtmlFromRtf(rtf);
      if (html) return html;
    }
  }

  return '';
}

function decodeUtf8(data: Uint8Array): string {
  const decoder = new TextDecoder('utf-8');
  return decoder.decode(data);
}

// --- Headers ---

function decodeHeaders(streams: CfbStream[]): MessageHeaders {
  const dateStr = getStringProperty(streams, '', PidTagMessageDeliveryTime);
  const messageId = getStringProperty(streams, '', PidTagInternetMessageId);
  const inReplyTo = getStringProperty(streams, '', PidTagInReplyToId);
  const importance = getLongProperty(streams, '', PidTagImportance);

  let importanceLevel: 'low' | 'normal' | 'high' = 'normal';
  if (importance === 0) importanceLevel = 'low';
  else if (importance === 2) importanceLevel = 'high';

  return {
    date: dateStr,
    messageId,
    inReplyTo,
    importance: importanceLevel,
  };
}

// --- Recipients ---

function decodeRecipients(streams: CfbStream[]): Recipient[] {
  const recipients: Recipient[] = [];

  // Find all recipient storages: __recip_version1.0_#XXXXXXXX
  const recipientPrefixes = new Set<string>();
  for (const s of streams) {
    const match = s.path.match(/__recip_version1\.0_#[0-9A-Fa-f]{8}/);
    if (match) {
      recipientPrefixes.add(match[0]);
    }
  }

  for (const prefix of recipientPrefixes) {
    const name = getStringProperty(streams, prefix, PidTagDisplayName);
    const email =
      getStringProperty(streams, prefix, PidTagSmtpAddress) ||
      getStringProperty(streams, prefix, PidTagEmailAddress);
    const recipType = getLongProperty(streams, prefix, PidTagRecipientType);

    let type: 'to' | 'cc' | 'bcc' = 'to';
    if (recipType === MAPI_CC) type = 'cc';
    else if (recipType === MAPI_BCC) type = 'bcc';

    recipients.push({ name, email, type });
  }

  return recipients;
}

// --- Attachments ---

function decodeAttachments(streams: CfbStream[]): Attachment[] {
  const attachments: Attachment[] = [];

  // Find all attachment storages: __attach_version1.0_#XXXXXXXX
  const attachPrefixes = new Set<string>();
  for (const s of streams) {
    const match = s.path.match(/__attach_version1\.0_#[0-9A-Fa-f]{8}/);
    if (match) {
      attachPrefixes.add(match[0]);
    }
  }

  for (const prefix of attachPrefixes) {
    const filename =
      getStringProperty(streams, prefix, PidTagAttachLongFilename) ||
      getStringProperty(streams, prefix, PidTagAttachFilename);
    const mimeType = getStringProperty(streams, prefix, PidTagAttachMimeTag);
    const contentId = getStringProperty(streams, prefix, PidTagAttachContentId);
    const content = getBinaryProperty(streams, prefix, PidTagAttachDataBinary);
    const renderingPos = getLongProperty(streams, prefix, PidTagRenderingPosition);

    const data = content ?? new Uint8Array(0);
    const isInline = contentId !== '' || (renderingPos !== null && renderingPos !== 0xffffffff);

    attachments.push({
      filename,
      mimeType: mimeType || 'application/octet-stream',
      contentId,
      content: data,
      size: data.length,
      isInline,
    });
  }

  return attachments;
}
