/**
 * Decompress MAPI LZFu-compressed RTF data.
 * Based on MS-OXRTFCP specification.
 * Uses only Uint8Array — no Node.js Buffer dependency.
 */

// Pre-defined dictionary from MS-OXRTFCP specification
// This must be exactly 207 bytes, padded to fill dictionary positions 0–206
const INIT_DICT_STR =
  '{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript ' +
  '\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par ' +
  '\\pard\\plain\\f0\\fs20\\b\\i\\ul\\ob\\strike\\scaps\\caps\\outline}';

const INIT_DICT_SIZE = 207;
const DICT_MAX_SIZE = 4096;

/**
 * Decompress LZFu compressed RTF from MAPI PidTagRtfCompressed property.
 * @param data - Raw compressed RTF bytes from PidTagRtfCompressed stream
 * @returns Decompressed RTF string, or null if invalid
 */
export function decompressRtf(data: Uint8Array): string | null {
  if (data.length < 16) return null;

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const compSize = view.getUint32(0, true);
  const rawSize = view.getUint32(4, true);
  const magic = view.getUint32(8, true);
  // const crc = view.getUint32(12, true);

  // Check magic
  // 0x75465a4c = "LZFu" (compressed)
  // 0x414e4f4d = "MONA" (uncompressed)
  if (magic === 0x414e4f4d) {
    // Uncompressed — just read raw bytes after header
    const raw = data.slice(16, 16 + rawSize);
    return decodeAscii(raw);
  }

  if (magic !== 0x75465a4c) return null;

  // Initialize dictionary with the pre-defined content
  const dict = new Uint8Array(DICT_MAX_SIZE);
  for (let i = 0; i < INIT_DICT_SIZE && i < INIT_DICT_STR.length; i++) {
    dict[i] = INIT_DICT_STR.charCodeAt(i);
  }
  let dictWritePos = INIT_DICT_SIZE;

  const output: number[] = [];
  let pos = 16; // Skip header
  const end = Math.min(data.length, compSize + 4); // compSize includes the 4-byte header field itself

  while (pos < end) {
    // Read control byte
    if (pos >= data.length) break;
    const control = data[pos++];

    for (let i = 0; i < 8 && pos < end; i++) {
      if ((control >> i) & 1) {
        // Reference: 2 bytes = dictionary offset + length
        if (pos + 1 >= data.length) break;
        const ref1 = data[pos++];
        const ref2 = data[pos++];
        const offset = (ref1 << 4) | (ref2 >> 4);
        const length = (ref2 & 0x0f) + 2;

        for (let j = 0; j < length; j++) {
          const ch = dict[(offset + j) % DICT_MAX_SIZE];
          output.push(ch);
          dict[dictWritePos] = ch;
          dictWritePos = (dictWritePos + 1) % DICT_MAX_SIZE;
        }
      } else {
        // Literal byte
        if (pos >= data.length) break;
        const ch = data[pos++];
        output.push(ch);
        dict[dictWritePos] = ch;
        dictWritePos = (dictWritePos + 1) % DICT_MAX_SIZE;
      }
    }
  }

  return decodeAscii(new Uint8Array(output));
}

/**
 * Extract HTML from RTF that was wrapped by Outlook using \fromhtml.
 * Outlook stores HTML inside RTF using {\*\htmltag NNN content} groups.
 */
export function extractHtmlFromRtf(rtf: string): string | null {
  if (!rtf.includes('\\fromhtml')) {
    return null;
  }
  return extractOutlookHtml(rtf);
}

/**
 * Extract HTML from Outlook's \fromhtml RTF wrapping.
 *
 * Outlook encodes HTML in RTF using three constructs:
 *   1. {\*\htmltag NNN <html_markup>}  — HTML tags (elements, attributes)
 *   2. \htmlrtf ... \htmlrtf0           — RTF-only rendering hints (SKIP these)
 *   3. Plain text between constructs    — actual HTML text content
 *
 * Strategy: walk the RTF linearly. Collect content from htmltag groups
 * and plain text regions. Skip \htmlrtf ... \htmlrtf0 blocks.
 */
function extractOutlookHtml(rtf: string): string | null {
  const html: string[] = [];
  // Start parsing from the first htmltag, skipping the RTF header/font table
  let i = rtf.indexOf('{\\*\\htmltag');
  if (i === -1) return null;
  let inHtmlRtf = false; // true when inside \htmlrtf (RTF-only, skip)

  while (i < rtf.length) {
    // Check for {\*\htmltag group — these are always HTML content regardless of htmlrtf state
    if (rtf.startsWith('{\\*\\htmltag', i)) {
      const content = readHtmlTagGroup(rtf, i);
      html.push(content.text);
      i = content.endPos;
      continue;
    }

    if (rtf[i] === '\\') {
      // Check for \htmlrtf toggle before processing as generic control word
      if (rtf.startsWith('\\htmlrtf', i)) {
        i += 8; // skip \htmlrtf
        if (i < rtf.length && rtf[i] === '0') {
          inHtmlRtf = false;
          i++;
        } else {
          inHtmlRtf = true;
        }
        // Skip delimiter space
        if (i < rtf.length && rtf[i] === ' ') i++;
        continue;
      }

      // If in htmlrtf mode, skip this control word
      if (inHtmlRtf) {
        i++;
        const result = skipControlWord(rtf, i);
        i = result.pos;
        continue;
      }

      // Outside htmlrtf: handle escapes and control words
      i++; // skip backslash
      if (i >= rtf.length) break;

      if (rtf[i] === '\'') {
        // Hex character escape: \'XX
        i++;
        if (i + 1 < rtf.length) {
          const code = parseInt(rtf.substring(i, i + 2), 16);
          if (!isNaN(code)) html.push(win1252ToChar(code));
          i += 2;
        }
      } else if (rtf[i] === '{' || rtf[i] === '}' || rtf[i] === '\\') {
        html.push(rtf[i++]);
      } else {
        // Skip other RTF control words
        const result = skipControlWord(rtf, i);
        i = result.pos;
      }
    } else if (rtf[i] === '{') {
      if (inHtmlRtf) {
        // Skip entire nested group while in htmlrtf mode
        i++;
        let depth = 1;
        while (i < rtf.length && depth > 0) {
          if (rtf[i] === '\\' && i + 1 < rtf.length && rtf.startsWith('\\htmlrtf', i)) {
            // Handle \htmlrtf inside nested braces
            i += 8;
            if (i < rtf.length && rtf[i] === '0') {
              inHtmlRtf = false;
              i++;
            }
            if (i < rtf.length && rtf[i] === ' ') i++;
            if (!inHtmlRtf) break; // exit brace-skipping, resume normal parsing
          } else if (rtf[i] === '{') {
            depth++;
            i++;
          } else if (rtf[i] === '}') {
            depth--;
            i++;
          } else {
            i++;
          }
        }
      } else {
        i++; // skip structural brace
      }
    } else if (rtf[i] === '}') {
      i++;
    } else if (rtf[i] === '\r' || rtf[i] === '\n') {
      i++; // RTF line wrapping
    } else {
      if (!inHtmlRtf) {
        html.push(rtf[i]);
      }
      i++;
    }
  }

  const result = html.join('');
  return result || null;
}

/**
 * Read a {\*\htmltag NNN content} group, decoding RTF escapes inside.
 */
function readHtmlTagGroup(rtf: string, start: number): { text: string; endPos: number } {
  // Skip past {\*\htmltag
  let pos = start + 11;

  // Skip tag number
  while (pos < rtf.length && rtf[pos] >= '0' && rtf[pos] <= '9') pos++;
  // Skip space delimiter
  if (pos < rtf.length && rtf[pos] === ' ') pos++;

  // Read content until matching }
  let depth = 1;
  let content = '';
  while (pos < rtf.length && depth > 0) {
    if (rtf[pos] === '{') {
      depth++;
      pos++;
    } else if (rtf[pos] === '}') {
      depth--;
      if (depth > 0) content += rtf[pos];
      pos++;
    } else if (rtf[pos] === '\\') {
      pos++;
      if (pos < rtf.length && rtf[pos] === '\'') {
        pos++;
        if (pos + 1 < rtf.length) {
          const code = parseInt(rtf.substring(pos, pos + 2), 16);
          if (!isNaN(code)) content += win1252ToChar(code);
          pos += 2;
        }
      } else if (pos < rtf.length && (rtf[pos] === '{' || rtf[pos] === '}' || rtf[pos] === '\\')) {
        content += rtf[pos++];
      } else {
        // Inside htmltag, \par means newline in the HTML source
        const result = skipControlWord(rtf, pos);
        if (result.word === 'par') content += '\n';
        else if (result.word === 'tab') content += '\t';
        pos = result.pos;
      }
    } else if (rtf[pos] === '\r' || rtf[pos] === '\n') {
      // RTF line breaks inside htmltag are not significant
      pos++;
    } else {
      content += rtf[pos++];
    }
  }

  return { text: content, endPos: pos };
}

/**
 * Skip an RTF control word and return the word name and new position.
 */
function skipControlWord(rtf: string, pos: number): { word: string; pos: number } {
  let word = '';
  while (pos < rtf.length && rtf[pos] >= 'a' && rtf[pos] <= 'z') {
    word += rtf[pos++];
  }
  // Skip optional numeric parameter
  if (pos < rtf.length && (rtf[pos] === '-' || (rtf[pos] >= '0' && rtf[pos] <= '9'))) {
    if (rtf[pos] === '-') pos++;
    while (pos < rtf.length && rtf[pos] >= '0' && rtf[pos] <= '9') pos++;
  }
  // Skip delimiter space
  if (pos < rtf.length && rtf[pos] === ' ') pos++;
  return { word, pos };
}

/**
 * Windows-1252 to Unicode mapping for bytes 0x80–0x9F.
 * These bytes are undefined in ISO-8859-1 but have characters in Windows-1252.
 */
const WIN1252_MAP: Record<number, number> = {
  0x80: 0x20AC, // €
  0x82: 0x201A, // ‚
  0x83: 0x0192, // ƒ
  0x84: 0x201E, // „
  0x85: 0x2026, // …
  0x86: 0x2020, // †
  0x87: 0x2021, // ‡
  0x88: 0x02C6, // ˆ
  0x89: 0x2030, // ‰
  0x8A: 0x0160, // Š
  0x8B: 0x2039, // ‹
  0x8C: 0x0152, // Œ
  0x8E: 0x017D, // Ž
  0x91: 0x2018, // '
  0x92: 0x2019, // '
  0x93: 0x201C, // "
  0x94: 0x201D, // "
  0x95: 0x2022, // •
  0x96: 0x2013, // –
  0x97: 0x2014, // —
  0x98: 0x02DC, // ˜
  0x99: 0x2122, // ™
  0x9A: 0x0161, // š
  0x9B: 0x203A, // ›
  0x9C: 0x0153, // œ
  0x9E: 0x017E, // ž
  0x9F: 0x0178, // Ÿ
};

/**
 * Decode a Windows-1252 byte value to a Unicode character.
 */
function win1252ToChar(code: number): string {
  if (code >= 0x80 && code <= 0x9F && WIN1252_MAP[code]) {
    return String.fromCharCode(WIN1252_MAP[code]);
  }
  return String.fromCharCode(code);
}

function decodeAscii(data: Uint8Array): string {
  const chars: string[] = [];
  for (let i = 0; i < data.length; i++) {
    chars.push(String.fromCharCode(data[i]));
  }
  return chars.join('');
}
