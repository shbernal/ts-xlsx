// The BIFF12 record framing every `.bin` part of an `.xlsb` package is built from ([MS-XLSB] 2.1.4).
//
// A part is a bare concatenation of records — no header, no index, no terminator — each framed as a
// variable-length type, a variable-length size, and that many payload bytes. Both prefixes are
// 7-bits-per-byte with the high bit meaning "one more byte follows": the type is 1–2 bytes (so the
// single-byte space is reserved for the hot records — a cell is one byte of framing), the size is 1–4.
//
// **This is the reader's hostile-input frontier.** The declared size is attacker-controlled, so it is
// never allowed to drive an allocation: a record's payload is handed out as a `subarray` *view* onto
// the part the inflate bound already materialised and capped, and a size that would run past the end
// of the part is rejected outright rather than clamped. A lying length therefore costs nothing and
// fails closed, and the whole stream's memory is exactly the part's own bytes.

import {XlsbParseError} from './errors.ts';

/** One framed BIFF12 record: its type, and a **view** onto its payload (never a copy). */
export interface BiffRecord {
  readonly type: number;
  readonly data: Uint8Array;
}

/**
 * Decode a `.bin` part into its record sequence, in order.
 *
 * Lazy: a caller that stops early (having found what it needs) never frames the rest of the part.
 *
 * @throws {XlsbParseError} if a record header is truncated or a record's declared size runs past the
 *   end of the part.
 */
export function* readRecords(part: Uint8Array): Generator<BiffRecord> {
  let offset = 0;
  while (offset < part.length) {
    const lowType = byteAt(part, offset++);
    let type = lowType & 0x7f;
    if ((lowType & 0x80) !== 0) type |= (byteAt(part, offset++) & 0x7f) << 7;

    // Up to four size bytes, least-significant seven bits first. The fourth byte's continuation bit
    // is specified as ignored, which the fixed bound expresses: the loop simply stops there.
    let size = 0;
    for (let index = 0; index < 4; index++) {
      const piece = byteAt(part, offset++);
      size |= (piece & 0x7f) << (7 * index);
      if ((piece & 0x80) === 0) break;
    }

    // The one check that makes a forged size harmless: compare it against what the part *actually*
    // holds, before anything is handed out. `part.length - offset` cannot be negative here — byteAt
    // has already proven every header byte was in range.
    if (size > part.length - offset) {
      throw new XlsbParseError(
        `BIFF12 record ${type} declares ${size} bytes but only ${part.length - offset} remain in the part`,
      );
    }
    yield {type, data: part.subarray(offset, offset + size)};
    offset += size;
  }
}

// Every header byte goes through here, so a part that ends mid-header fails closed rather than
// folding `undefined` into the arithmetic (which `noUncheckedIndexedAccess` would otherwise let
// through as a silent NaN-shaped type or size).
function byteAt(part: Uint8Array, index: number): number {
  const value = part[index];
  if (value === undefined) throw new XlsbParseError('truncated BIFF12 record header');
  return value;
}
