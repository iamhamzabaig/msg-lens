/** Represents a successfully parsed .msg email */
export interface ParsedMessage {
  /** Email subject line */
  subject: string;
  /** Sender display name */
  senderName: string;
  /** Sender email address */
  senderEmail: string;
  /** Recipients (To) */
  recipients: Recipient[];
  /** CC recipients */
  ccRecipients: Recipient[];
  /** BCC recipients */
  bccRecipients: Recipient[];
  /** Plain text body (if available) */
  bodyText: string;
  /** Sanitized HTML body with cid: references resolved to base64 */
  bodyHtml: string;
  /** Email headers (date, message-id, etc.) */
  headers: MessageHeaders;
  /** File attachments */
  attachments: Attachment[];
}

/** Email recipient */
export interface Recipient {
  name: string;
  email: string;
  type: 'to' | 'cc' | 'bcc';
}

/** Common email headers */
export interface MessageHeaders {
  /** Date as ISO 8601 string */
  date: string;
  /** Date as JS Date object, null if unavailable */
  dateObject: Date | null;
  messageId: string;
  inReplyTo: string;
  importance: 'low' | 'normal' | 'high';
}

/** File attachment or embedded image */
export interface Attachment {
  /** Original filename */
  filename: string;
  /** MIME type */
  mimeType: string;
  /** Content-ID for inline/embedded images */
  contentId: string;
  /** Raw attachment bytes */
  content: Uint8Array;
  /** Size in bytes */
  size: number;
  /** Whether this is an inline/embedded attachment */
  isInline: boolean;
  /** Parsed embedded .msg message (when attachment is an embedded email) */
  embeddedMessage?: ParsedMessage;
}

/** Result wrapper — either success or error */
export type ParseResult =
  | { success: true; message: ParsedMessage }
  | { success: false; error: ParseError };

/** Typed parse error */
export interface ParseError {
  code: ParseErrorCode;
  message: string;
}

export type ParseErrorCode =
  | 'INVALID_CFB'
  | 'INVALID_EML'
  | 'MISSING_PROPERTIES'
  | 'MALFORMED_MAPI'
  | 'MALFORMED_MIME'
  | 'SANITIZATION_FAILED'
  | 'UNKNOWN_ERROR';
