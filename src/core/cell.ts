// A single cell: a typed value at a fixed 1-based `(row, col)` position.
//
// Position is immutable and numeric — the corpus locks `col`/`row` as 1-based
// *numbers* (legacy shipped a type declaration calling them strings, which broke
// strict consumers). The value is the only mutable state here; assigning it routes
// through the value model so the cell's `type` is always consistent with what it holds.

import {encodeAddress} from './address.ts';
import type {Alignment, Border, Fill, Font, Protection} from './style.ts';
import {type CellValue, coerceCellValue, detectValueType, type ValueType} from './value.ts';

export class Cell {
  /** 1-based row index. */
  readonly row: number;
  /** 1-based column index. */
  readonly col: number;

  #value: CellValue = null;
  #fill: Fill | undefined;
  #numFmt: string | undefined;
  #font: Partial<Font> | undefined;
  #border: Border | undefined;
  #alignment: Alignment | undefined;
  #protection: Protection | undefined;
  #quotePrefix: boolean | undefined;
  #namedStyleId: number | undefined;
  #note: string | undefined;

  constructor(row: number, col: number) {
    if (!Number.isInteger(row) || row < 1) {
      throw new RangeError(`cell row ${row} is out of bounds — rows start at 1`);
    }
    if (!Number.isInteger(col) || col < 1) {
      throw new RangeError(`cell column ${col} is out of bounds — columns start at 1`);
    }
    this.row = row;
    this.col = col;
  }

  /** Canonical A1 address of this cell (`"B3"`). */
  get address(): string {
    return encodeAddress(this.col, this.row);
  }

  /** The cell's value; `null` when empty. Assigning `undefined` clears it. */
  get value(): CellValue {
    return this.#value;
  }

  set value(value: CellValue | undefined) {
    this.#value = coerceCellValue(value);
  }

  /** The observable {@link ValueType} of the current value. */
  get type(): ValueType {
    return detectValueType(this.#value);
  }

  /**
   * The cell's background fill, or `undefined` when it has none. Each cell owns its
   * own fill; assigning one never aliases a neighbour's style, so a fill set on one
   * cell cannot bleed onto its row, column, or sheet siblings.
   */
  get fill(): Fill | undefined {
    return this.#fill;
  }

  set fill(fill: Fill | undefined) {
    this.#fill = fill;
  }

  /**
   * The cell's number-format code (`"0.00%"`, a custom accounting format, …), or
   * `undefined` for the General format. Stored verbatim: the invariant form Excel
   * persists — `.` decimal, `,` grouping, `/` date separator — is neither localized
   * nor rewritten, so the code round-trips character-for-character. Like {@link fill},
   * each cell owns its own code; a cell that also carries a column-level format keeps
   * both, so overriding one facet never drops the other.
   */
  get numFmt(): string | undefined {
    return this.#numFmt;
  }

  set numFmt(numFmt: string | undefined) {
    this.#numFmt = numFmt;
  }

  /**
   * The cell's font — bold/italic/underline, size, colour, typeface — as a partial set
   * of the facets that differ from the default (only the facets actually set are carried,
   * exactly as OOXML stores them). `undefined` means the cell uses the workbook default
   * font. Like {@link fill} and {@link numFmt}, each cell owns its own font object, so a
   * font set on one cell never aliases or bleeds onto its row, column, or sheet siblings.
   */
  get font(): Partial<Font> | undefined {
    return this.#font;
  }

  set font(font: Partial<Font> | undefined) {
    this.#font = font;
  }

  /**
   * The cell's border — the line style and colour of each side — or `undefined` when the
   * cell has none. An absent edge within a border means that side is unbordered, so reading
   * a cell never fabricates a border it does not have. Like {@link fill}, {@link numFmt}, and
   * {@link font}, each cell owns its own border object, so a border set on one cell never
   * aliases or bleeds onto its row, column, or sheet siblings.
   */
  get border(): Border | undefined {
    return this.#border;
  }

  set border(border: Border | undefined) {
    this.#border = border;
  }

  /**
   * The cell's alignment — how its content sits within the cell, plus the wrap/shrink flags —
   * or `undefined` when it uses the defaults. The boolean flags are off unless explicitly set,
   * so a cell that never enabled wrapping never reads back wrapped. Like {@link fill},
   * {@link numFmt}, {@link font}, and {@link border}, each cell owns its own alignment object, so
   * an alignment set on one cell never aliases or bleeds onto its row, column, or sheet siblings.
   */
  get alignment(): Alignment | undefined {
    return this.#alignment;
  }

  set alignment(alignment: Alignment | undefined) {
    this.#alignment = alignment;
  }

  /**
   * The cell's protection — its locked/hidden flags, enforced only once the sheet is protected —
   * or `undefined` when the cell carries neither. `locked` defaults to on in OOXML, so a cell
   * that never touched protection is implicitly locked and reads back as `undefined`, not as
   * `{locked: true}`; the flag only becomes explicit when a cell is unlocked. Like {@link fill},
   * {@link numFmt}, {@link font}, {@link border}, and {@link alignment}, each cell owns its own
   * protection object, so protection set on one cell never aliases or bleeds onto its siblings.
   */
  get protection(): Protection | undefined {
    return this.#protection;
  }

  set protection(protection: Protection | undefined) {
    this.#protection = protection;
  }

  /**
   * The quote-prefix flag: when set, a spreadsheet stores the cell's content as literal text even
   * when it looks like a formula or number, and shows a leading apostrophe in the formula bar without
   * that apostrophe being part of the stored value. `undefined` (or `false`) when unset. It is a
   * cell-format flag — an attribute on the cell's `xf` record — so it composes independently of the
   * value and, like the other style facets, never bleeds onto sibling cells.
   */
  get quotePrefix(): boolean | undefined {
    return this.#quotePrefix;
  }

  set quotePrefix(quotePrefix: boolean | undefined) {
    this.#quotePrefix = quotePrefix;
  }

  /**
   * The index of the {@link Workbook.namedStyles named cell style} this cell links to (its OOXML
   * `xfId`), or `undefined` when the cell references no named style beyond the default. The cell
   * inherits any facet its own direct format leaves unset from that named style; the reader resolves
   * the effective look onto the cell's own facets, and this link is preserved so a round-trip keeps
   * the cell tied to its named style rather than flattening it away.
   */
  get namedStyleId(): number | undefined {
    return this.#namedStyleId;
  }

  set namedStyleId(namedStyleId: number | undefined) {
    this.#namedStyleId = namedStyleId;
  }

  /**
   * The cell's note (comment) as plain text, or `undefined` when it carries none. A note is
   * metadata anchored to the cell, independent of its value: a cell can hold a note while empty,
   * and clearing the value leaves the note intact. Like the style facets, each cell owns its own
   * note, so a note set on one cell never bleeds onto its siblings, and a structural edit that
   * shifts the cell carries the note along to its new position.
   */
  get note(): string | undefined {
    return this.#note;
  }

  set note(note: string | undefined) {
    this.#note = note;
  }
}

// Copy a cell's value and every style facet onto another cell. Used when a structural edit shifts a
// cell to a new position: `Cell` fixes its `(row, col)` at construction, so a shifted cell is a fresh
// cell at the new coordinates carrying the original's content. Facet objects are passed by reference —
// safe under the copy-on-write style model (facet setters replace, never mutate), so the source and
// its shifted copy never alias each other's style through a shared object.
export function copyCellContent(source: Cell, target: Cell): void {
  target.value = source.value;
  target.fill = source.fill;
  target.numFmt = source.numFmt;
  target.font = source.font;
  target.border = source.border;
  target.alignment = source.alignment;
  target.protection = source.protection;
  target.note = source.note;
}
