// A single cell: a typed value at a fixed 1-based `(row, col)` position.
//
// Position is immutable and numeric — the corpus locks `col`/`row` as 1-based
// *numbers* (legacy shipped a type declaration calling them strings, which broke
// strict consumers). The value is the only mutable state here; assigning it routes
// through the value model so the cell's `type` is always consistent with what it holds.

import {encodeAddress} from './address.ts';
import type {Alignment, Border, CellStyle, Fill, Font, Protection} from './style.ts';
import {type CellValue, coerceCellValue, detectValueType, type ValueType} from './value.ts';
import type {CellModel} from './worksheet.ts';

/**
 * A single cell owns its value and every style facet outright. Each facet below — fill, number format,
 * font, border, alignment, protection, quote-prefix, and note — is held in the cell's own field and
 * *replaced* (never mutated in place) by its setter, so a facet set on one cell never aliases or bleeds
 * onto its row, column, or sheet siblings. Each facet's own doc covers only what is specific to it.
 */
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

  /** The cell's background fill, or `undefined` when it has none. */
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
   * nor rewritten, so the code round-trips character-for-character. A cell that also carries
   * a column-level format keeps both, so overriding one facet never drops the other.
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
   * exactly as OOXML stores them). `undefined` means the cell uses the workbook default font.
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
   * a cell never fabricates a border it does not have.
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
   * so a cell that never enabled wrapping never reads back wrapped.
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
   * `{locked: true}`; the flag only becomes explicit when a cell is unlocked.
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
   * value.
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
   * and clearing the value leaves the note intact. A structural edit that shifts the cell carries the
   * note along to its new position.
   */
  get note(): string | undefined {
    return this.#note;
  }

  set note(note: string | undefined) {
    this.#note = note;
  }
}

// Lay each present style facet of `style` onto `cell`, leaving facets it omits untouched. This is the
// one place the six-facet {@link CellStyle} tuple is walked onto a cell, shared by every path that
// applies a style — a table column's format, a resolved read xf, a model assignment — so the facet
// list lives in a single spot and no path can silently forget one. Facet objects are assigned by
// reference, safe under the copy-on-write style model (setters replace, never mutate in place).
export function applyCellStyle(cell: Cell, style: Readonly<CellStyle>): void {
  if (style.fill !== undefined) cell.fill = style.fill;
  if (style.numFmt !== undefined) cell.numFmt = style.numFmt;
  if (style.font !== undefined) cell.font = style.font;
  if (style.border !== undefined) cell.border = style.border;
  if (style.alignment !== undefined) cell.alignment = style.alignment;
  if (style.protection !== undefined) cell.protection = style.protection;
}

// Copy a cell's value and every style facet onto a target cell. The source is a {@link CellModel},
// which a live {@link Cell} structurally satisfies, so this one primitive serves both directions that
// load content into a cell: a structural edit shifting a cell to fresh coordinates (`Cell` fixes its
// `(row, col)` at construction, so the shifted cell is a new cell carrying the original's content) and
// assigning a {@link WorksheetModel} onto a sheet. Position is never copied — the target keeps its own
// `(row, col)`. The style facets go through {@link applyCellStyle} (targets are always fresh cells, so
// its skip-if-absent is equivalent to a full copy here). Paired with {@link cellToModel} (the read
// direction); a facet cellToModel emits but applyCellStyle omits (or the reverse) would silently drop
// on a model round-trip — the historical merge-loss failure the CellStyle tuple now guards by type.
export function copyCellContent(source: CellModel, target: Cell): void {
  target.value = source.value;
  applyCellStyle(target, source);
  target.note = source.note;
}

// Snapshot a cell's position and content as a {@link CellModel} — the read direction paired with
// {@link copyCellContent}'s write. Emits exactly the facets {@link applyCellStyle} consumes, so a
// `dst.model = src.model` round-trip carries every one; a facet listed here but not there (or the
// reverse) is precisely the silent merge-loss the {@link WorksheetModel} contract guards against.
export function cellToModel(cell: Cell): CellModel {
  return {
    row: cell.row,
    col: cell.col,
    value: cell.value,
    fill: cell.fill,
    numFmt: cell.numFmt,
    font: cell.font,
    border: cell.border,
    alignment: cell.alignment,
    protection: cell.protection,
    note: cell.note,
  };
}
