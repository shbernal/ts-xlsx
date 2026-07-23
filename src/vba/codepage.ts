// The VBA project stores module names and source as MBCS bytes in the project code page
// (`PROJECTCODEPAGE`, [MS-OVBA] 2.3.4.2). To recover text we decode those bytes with the matching
// encoding rather than assuming latin1 — a CJK or Cyrillic project would otherwise mojibake.

import {VbaAuthorError, VbaParseError} from './errors.ts';

// `TextDecoder` is a global value here, not exposed as a type by the configured libs; alias its
// instance type so decoders can be passed around with a name.
export type Decoder = InstanceType<typeof TextDecoder>;

/** Encodes text back to the project code page's MBCS bytes; throws on a character it cannot represent. */
export type Encoder = (text: string) => Uint8Array;

// Windows code-page numbers → the WHATWG encoding label `TextDecoder` understands. Node's built-in
// ICU covers all of these. Only the code pages a VBA project realistically declares are listed; an
// unlisted one falls back to windows-1252 (the overwhelmingly common Western-European default) so a
// rare project still decodes readably instead of throwing.
const CODEPAGE_LABEL = new Map<number, string>([
  [1250, 'windows-1250'],
  [1251, 'windows-1251'],
  [1252, 'windows-1252'],
  [1253, 'windows-1253'],
  [1254, 'windows-1254'],
  [1255, 'windows-1255'],
  [1256, 'windows-1256'],
  [1257, 'windows-1257'],
  [1258, 'windows-1258'],
  [874, 'windows-874'],
  [932, 'shift_jis'],
  [936, 'gbk'],
  [949, 'euc-kr'],
  [950, 'big5'],
  [10000, 'macintosh'],
  [20866, 'koi8-r'],
  [21866, 'koi8-u'],
  [28591, 'iso-8859-1'],
  [28592, 'iso-8859-2'],
  [65001, 'utf-8'],
]);

/**
 * A `TextDecoder` for the given VBA project code page. Non-fatal (malformed bytes become U+FFFD rather
 * than throwing) because recovered source is for reading, not re-encoding — a stray byte must not sink
 * the whole extraction. An unknown code page falls back to windows-1252.
 */
export function decoderForCodePage(codePage: number): Decoder {
  const label = CODEPAGE_LABEL.get(codePage) ?? 'windows-1252';
  try {
    return new TextDecoder(label, {fatal: false});
  } catch {
    // A runtime whose ICU lacks the label (a minimal build) still gets usable Western-European text
    // rather than a hard failure on an otherwise-valid project.
    try {
      return new TextDecoder('windows-1252', {fatal: false});
    } catch (cause) {
      throw new VbaParseError('runtime TextDecoder cannot decode the VBA project code page', {
        cause,
      });
    }
  }
}

/**
 * The write-side inverse of {@link decoderForCodePage}: an encoder for a project's code page, built by
 * decoding every byte 0x00..0xFF through the matching decoder and inverting the map. This is exact for
 * the single-byte code pages a VBA project realistically declares. Under a multi-byte code page (Shift
 * JIS, GBK, …) the high bytes are lead bytes, not standalone characters, so only ASCII inverts and any
 * non-ASCII character is rejected — fail-closed rather than emit a corrupt byte. A character the chosen
 * code page cannot represent throws {@link VbaAuthorError}.
 */
export function encoderForCodePage(codePage: number): Encoder {
  const decoder = decoderForCodePage(codePage);
  const charToByte = new Map<number, number>();
  const one = new Uint8Array(1);
  for (let b = 0x00; b <= 0xff; b++) {
    one[0] = b;
    const decoded = decoder.decode(one);
    if (decoded.length !== 1) continue; // a lead byte / incomplete sequence under a multi-byte code page
    const code = decoded.charCodeAt(0);
    if (code !== 0xfffd && !charToByte.has(code)) charToByte.set(code, b);
  }
  return (text: string): Uint8Array => {
    const out = new Uint8Array(text.length);
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      const byte = charToByte.get(code);
      if (byte === undefined) {
        throw new VbaAuthorError(
          `character U+${code.toString(16).toUpperCase().padStart(4, '0')} is not representable in code page ${codePage}`,
        );
      }
      out[i] = byte;
    }
    return out;
  };
}
