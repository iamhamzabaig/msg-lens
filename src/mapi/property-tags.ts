/**
 * MAPI property tag constants used in .msg files.
 * Format: 0xPPPPTTTT where PPPP = property ID, TTTT = property type.
 */

// Property types
export const PT_UNICODE = 0x001f;
export const PT_STRING8 = 0x001e;
export const PT_BINARY = 0x0102;
export const PT_LONG = 0x0003;
export const PT_BOOLEAN = 0x000b;
export const PT_SYSTIME = 0x0040;

// Message properties (property IDs)
export const PidTagSubject = 0x0037;
export const PidTagSenderName = 0x0c1a;
export const PidTagSenderEmailAddress = 0x0c1f;
export const PidTagSenderSmtpAddress = 0x5d01;
export const PidTagBody = 0x1000;
export const PidTagHtml = 0x1013;
export const PidTagRtfCompressed = 0x1009;
export const PidTagMessageDeliveryTime = 0x0e06;
export const PidTagInternetMessageId = 0x1035;
export const PidTagInReplyToId = 0x1042;
export const PidTagImportance = 0x0017;
export const PidTagDisplayTo = 0x0e04;
export const PidTagDisplayCc = 0x0e03;
export const PidTagDisplayBcc = 0x0e02;
export const PidTagTransportMessageHeaders = 0x007d;

// Recipient properties
export const PidTagDisplayName = 0x3001;
export const PidTagEmailAddress = 0x3003;
export const PidTagSmtpAddress = 0x39fe;
export const PidTagRecipientType = 0x0c15;

// Attachment properties
export const PidTagAttachFilename = 0x3704;
export const PidTagAttachLongFilename = 0x3707;
export const PidTagAttachMimeTag = 0x370e;
export const PidTagAttachContentId = 0x3712;
export const PidTagAttachDataBinary = 0x3701;
export const PidTagAttachMethod = 0x3705;
export const PidTagAttachFlags = 0x3714;
export const PidTagAttachRendering = 0x3709;
export const PidTagRenderingPosition = 0x370b;

// Recipient types
export const MAPI_TO = 1;
export const MAPI_CC = 2;
export const MAPI_BCC = 3;
