// The [MS-XLSB] primitive types a record payload is built from, behind one bounds-checked cursor.
//
// Binary parsing is where a reader most easily goes wrong in two directions at once: silently reading
// past the end of a buffer, and allocating on a length the file (not the reader) chose. `RecordReader`
// closes both. Every read goes through a single `#take`, so overrunning the record is impossible
// rather than merely unlikely; and a length-prefixed string checks its byte count against what the
// record actually holds *before* a single character is materialised, so a forged `cch` costs one
// comparison instead of gigabytes.
//
// Structures decoded here — RkNumber, XLWideString, BrtColor, Cell, UncheckedRfX, BErr — are shared
// across the workbook, worksheet, styles, and shared-string parsers; nothing part-specific lives here.

import type {Color} from '../../core/style.ts';
import {type ErrorCode, isErrorCode} from '../../core/value.ts';
import {XlsbParseError} from './errors.ts';

/** The `Cell` structure ([MS-XLSB] 2.5.10) every cell record opens with. */
export interface CellHeader {
  /** Zero-based column index. */
  readonly column: number;
  /** Zero-based index into the style sheet's cell-XF collection. */
  readonly styleIndex: number;
}

/** An `UncheckedRfX` ([MS-XLSB] 2.5.155) cell range — all four bounds zero-based and inclusive. */
export interface RangeBounds {
  readonly rowFirst: number;
  readonly rowLast: number;
  readonly colFirst: number;
  readonly colLast: number;
}

// A length-prefixed string is decoded in code-unit batches rather than one `String.fromCharCode` call
// per character (quadratic concatenation) or one spread of every unit (which blows the argument limit
// on a long string). 4096 is comfortably under every engine's limit and makes the batching invisible.
const CHARS_PER_BATCH = 4096;

// `XLNullableWideString` marks "no string" with a character count of 0xFFFFFFFF rather than 0 — an
// empty string and an absent one are different values (a sheet's relationship id is nullable; its
// name is not).
const NULL_STRING_LENGTH = 0xffffffff;

// One reusable 8-byte window for reassembling an RkNumber's truncated double. The alternative — a
// fresh ArrayBuffer per RK cell — would allocate once per numeric cell in the workbook, on the single
// hottest path in the reader. Safe to share: the write and the read below are one synchronous pair.
const rkScratch = new DataView(new ArrayBuffer(8));

/**
 * A bounds-checked cursor over one record's payload.
 *
 * Each accessor advances the cursor by exactly the bytes it consumed, so a record is decoded by
 * naming its fields in order. Reading past the payload throws {@link XlsbParseError} — a record that
 * is shorter than its own definition is a malformed file, not a case to guess through.
 */
export class RecordReader {
  readonly #data: Uint8Array;
  readonly #view: DataView;
  #offset = 0;

  constructor(data: Uint8Array) {
    this.#data = data;
    this.#view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  }

  /** Bytes left in the record. */
  get remaining(): number {
    return this.#data.length - this.#offset;
  }

  /** Whether the whole payload has been consumed. */
  get done(): boolean {
    return this.remaining <= 0;
  }

  /** Advance past `count` bytes without decoding them — a reserved or unmodelled field. */
  skip(count: number): void {
    this.#take(count);
  }

  /**
   * The next `count` bytes as a **view**, for a field whose own decoding happens elsewhere — a formula
   * token stream, whose meaning depends on workbook tables this record knows nothing about. A view
   * rather than a copy for the same reason a record's payload is one: the declared length comes from
   * the file, so it must bound a read, never an allocation.
   */
  bytes(count: number): Uint8Array {
    return this.#data.subarray(this.#take(count), this.#offset);
  }

  u8(): number {
    return this.#view.getUint8(this.#take(1));
  }

  u16(): number {
    return this.#view.getUint16(this.#take(2), true);
  }

  i16(): number {
    return this.#view.getInt16(this.#take(2), true);
  }

  u32(): number {
    return this.#view.getUint32(this.#take(4), true);
  }

  i32(): number {
    return this.#view.getInt32(this.#take(4), true);
  }

  /** An `Xnum` ([MS-XLSB] 2.5.172): a little-endian IEEE-754 double. */
  f64(): number {
    return this.#view.getFloat64(this.#take(8), true);
  }

  /**
   * An `RkNumber` ([MS-XLSB] 2.5.122): a number packed into 32 bits. Two flag bits steal the low end
   * of the word — `fInt` says the remaining 30 bits are a signed integer rather than the *high* 30
   * bits of a double whose low 34 bits are zero, and `fX100` says the result was scaled up by 100 to
   * keep two decimal places in the integer form. It exists because most real spreadsheet numbers are
   * small integers or two-decimal currency, and this stores them in half the bytes of a double.
   */
  rk(): number {
    // Read signed: the `fInt` branch needs an arithmetic shift to sign-extend a negative 30-bit
    // integer, which only works on a value JavaScript already considers negative.
    const raw = this.i32();
    const value = (raw & 0b10) !== 0 ? raw >> 2 : truncatedDouble(raw & ~0b11);
    return (raw & 0b01) !== 0 ? value / 100 : value;
  }

  /** An `XLWideString` ([MS-XLSB] 2.5.169): a 4-byte character count then that many UTF-16LE units. */
  wideString(): string {
    return this.#characters(this.u32());
  }

  /**
   * A UTF-16 string whose character count is 16-bit rather than 32-bit — the form used *inside* a
   * formula token stream (`PtgStr`, and the string elements of an array constant), where a 4-byte
   * count on every literal would be pure overhead.
   */
  shortString(): string {
    return this.#characters(this.u16());
  }

  /** An `XLNullableWideString` ([MS-XLSB] 2.5.166): an {@link wideString} that can also be absent. */
  nullableWideString(): string | undefined {
    const length = this.u32();
    return length === NULL_STRING_LENGTH ? undefined : this.#characters(length);
  }

  /**
   * A `RichStr` ([MS-XLSB] 2.5.124): a string that may carry per-run formatting and phonetic guides.
   * Only the text is returned — the run and phonetic tails are left unread, which is safe because the
   * record's framing (not this cursor) bounds where the payload ends.
   */
  richString(): string {
    this.skip(1); // fRichStr / fExtStr flags: which optional tails follow the text.
    return this.wideString();
  }

  /** The `Cell` structure ([MS-XLSB] 2.5.10) that opens every cell record. */
  cell(): CellHeader {
    const column = this.u32();
    // The style index shares its word with a phonetic-display flag in the high byte.
    return {column, styleIndex: this.u32() & 0x00ffffff};
  }

  /** An `UncheckedRfX` ([MS-XLSB] 2.5.155): four zero-based, inclusive range bounds. */
  range(): RangeBounds {
    return {
      rowFirst: this.u32(),
      rowLast: this.u32(),
      colFirst: this.u32(),
      colLast: this.u32(),
    };
  }

  /**
   * A `BrtColor` ([MS-XLSB] 2.4.337), mapped onto the model's {@link Color}.
   *
   * The four encodings are mutually exclusive and the type tag picks which of the payload's fields
   * carry meaning; the rest are explicitly undefined. An *automatic* colour (type 0) names nothing at
   * all, and reads back as no colour — the same absence the XML reader produces for `<color auto="1"/>`,
   * so a cell whose font colour was never set does not gain one on read.
   */
  color(): Color | undefined {
    const flags = this.u8();
    const index = this.u8();
    const tintAndShade = this.i16();
    const red = this.u8();
    const green = this.u8();
    const blue = this.u8();
    const alpha = this.u8();
    const base = colorByType(flags >> 1, index, alpha, red, green, blue);
    if (base === undefined) return undefined;
    // Tint is stored as a fraction of the signed 16-bit range, where the extreme values mean 100%
    // lightening/darkening; the model carries it as OOXML does, in [-1, 1].
    if (tintAndShade === 0) return base;
    return {...base, tint: Math.max(-1, Math.min(1, tintAndShade / 0x7fff))};
  }

  #characters(count: number): string {
    // Check the byte count against the record *before* building anything: this is the guard that
    // makes a forged character count a cheap failure rather than an allocation the file chose.
    const start = this.#take(count * 2);
    let text = '';
    let batch: number[] = [];
    for (let index = 0; index < count; index++) {
      batch.push(this.#view.getUint16(start + index * 2, true));
      if (batch.length === CHARS_PER_BATCH) {
        text += String.fromCharCode(...batch);
        batch = [];
      }
    }
    return batch.length > 0 ? text + String.fromCharCode(...batch) : text;
  }

  // The single choke point every read passes through. Returns the offset the caller may read from,
  // having proven that `count` bytes are there.
  #take(count: number): number {
    if (count > this.remaining) {
      throw new XlsbParseError(
        `BIFF12 record field needs ${count} bytes but only ${this.remaining} remain in the record`,
      );
    }
    const start = this.#offset;
    this.#offset += count;
    return start;
  }
}

/** A `BErr` ([MS-XLSB] 2.5.98.2) error code, as the model's error string. */
export function errorCodeFor(code: number): ErrorCode | undefined {
  const text = BERR_CODES.get(code);
  return text !== undefined && isErrorCode(text) ? text : undefined;
}

const BERR_CODES: ReadonlyMap<number, string> = new Map([
  [0x00, '#NULL!'],
  [0x07, '#DIV/0!'],
  [0x0f, '#VALUE!'],
  [0x17, '#REF!'],
  [0x1d, '#NAME?'],
  [0x24, '#NUM!'],
  [0x2a, '#N/A'],
  [0x2b, '#GETTING_DATA'],
]);

// Rebuild the double whose *high* 32 bits are `high` and whose low 34 bits [MS-XLSB] guarantees to be
// zero. Written big-endian so the given word lands in the high half regardless of host endianness.
function truncatedDouble(high: number): number {
  rkScratch.setInt32(0, high, false);
  rkScratch.setUint32(4, 0, false);
  return rkScratch.getFloat64(0, false);
}

// The `xColorType` tag ([MS-XLSB] 2.4.337) selects which of a BrtColor's fields carry the colour.
const COLOR_TYPE_INDEXED = 1;
const COLOR_TYPE_RGB = 2;
const COLOR_TYPE_THEME = 3;

function colorByType(
  type: number,
  index: number,
  alpha: number,
  red: number,
  green: number,
  blue: number,
): Color | undefined {
  switch (type) {
    case COLOR_TYPE_INDEXED:
      return {indexed: index};
    case COLOR_TYPE_RGB:
      return {argb: hexByte(alpha) + hexByte(red) + hexByte(green) + hexByte(blue)};
    case COLOR_TYPE_THEME:
      // The binary theme index and OOXML's `theme="…"` attribute both count the `clrScheme`
      // subelements in declaration order (dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink), so the
      // index carries across unchanged.
      return {theme: index};
    default:
      return undefined;
  }
}

function hexByte(value: number): string {
  return value.toString(16).toUpperCase().padStart(2, '0');
}
