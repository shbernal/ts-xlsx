// Cell-gathering state machine for the worksheet-body reader. It owns the still-textual pieces of the
// one `<c>` currently being read — its address/type/style, formula, cached value, inline text, and
// rich-text runs — plus the shared-formula master map that spans the whole sheet. Each `<c>` resets
// the per-cell state on {@link beginCell} and commits it on {@link finalize}, so this class is the
// single owner of "what has this cell gathered so far" — to a cell what {@link RunAccumulator} is to a
// rich string. Value *decoding* stays in `cell-value.ts`; this class only gathers the raw pieces.

import {decodeAddress, encodeAddress} from '../../core/address.ts';
import {applyCellStyle, type Cell} from '../../core/cell.ts';
import {translateFormula, unmangleFunctions} from '../../core/formula.ts';
import type {CellValue, DataTableFormulaValue, SharedFormulaValue} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {
  decodeCellContent,
  decodeFormulaResult,
  type RawCell,
  type SharedString,
} from './cell-value.ts';
import type {XfStyle} from './read-styles.ts';
import {RunAccumulator} from './rich-runs.ts';
import {boolPresent, type XmlAttributes} from './xml-read.ts';

// The declaration attributes of a `<f t="dataTable">`, held from the `<f>` open until the cell
// finalises. A data-table formula is preserved by declaration, not evaluated.
interface DataTableDeclaration {
  ref: string;
  dt2D: string | undefined;
  dtr: string | undefined;
  r1: string | undefined;
  r2: string | undefined;
}

export class CellAccumulator {
  #ref = '';
  #type = '';
  #style = -1;
  #col = -1;
  #row = -1;
  #formula = '';
  // Shared-formula bookkeeping. A master `<f t="shared" ref si>TEXT</f>` seeds the group; every clone
  // `<f t="shared" si/>` in the sheet references it by `si` and carries no text of its own.
  #formulaShared = false;
  #formulaSi = -1;
  #sharedClone = false;
  #dataTable: DataTableDeclaration | null = null;
  #valueText = '';
  #inlineText = '';
  #hasFormula = false;
  #hasValue = false;
  readonly #runs = new RunAccumulator();
  // Masters always precede their clones (Excel keeps the master top-left), so a clone resolves against
  // a map filled as the sheet streams: the master's formula translated to the clone's position.
  readonly #masters = new Map<number, {formula: string; col: number; row: number}>();

  /** This cell's `<c r>` address (`"B3"`), or '' when it carried none. */
  get ref(): string {
    return this.#ref;
  }

  /** This cell's own `<c s>` style index, or -1 when it carries none. */
  get styleIndex(): number {
    return this.#style;
  }

  /** This cell's 1-based column, or -1 when its address was absent or unparseable. */
  get col(): number {
    return this.#col;
  }

  /** The rich-text run accumulator, driven by the surrounding parser's `<r>`/`<rPr>` handling. */
  get runs(): RunAccumulator {
    return this.#runs;
  }

  // Begin a new `<c>`: record its address/type/style and clear every per-cell gathered field so the
  // last cell's formula, value, runs, or shared/data-table declaration cannot bleed into this one.
  beginCell(attrs: XmlAttributes): void {
    this.#ref = attrs.r ?? '';
    this.#type = attrs.t ?? '';
    this.#style = attrs.s !== undefined ? Number(attrs.s) : -1;
    this.#col = this.#ref === '' ? -1 : (decodeAddress(this.#ref).col ?? -1);
    this.#row = this.#ref === '' ? -1 : (decodeAddress(this.#ref).row ?? -1);
    this.#formula = '';
    this.#valueText = '';
    this.#inlineText = '';
    this.#runs.reset();
    this.#hasFormula = false;
    this.#hasValue = false;
    this.#formulaShared = false;
    this.#formulaSi = -1;
    this.#sharedClone = false;
    this.#dataTable = null;
  }

  // Begin an `<f>`: record its shared-formula grouping and any data-table declaration. A self-closing
  // `<f t="shared" si/>` is a clone — it fires no close and carries no text — so mark it here to
  // resolve against its master when the cell finalises.
  beginFormula(attrs: XmlAttributes, selfClosing: boolean): void {
    this.#formulaShared = attrs.t === 'shared';
    this.#formulaSi = attrs.si !== undefined ? Number(attrs.si) : -1;
    if (selfClosing && this.#formulaShared) this.#sharedClone = true;
    if (attrs.t === 'dataTable' && attrs.ref !== undefined) {
      this.#dataTable = {
        ref: attrs.ref,
        dt2D: attrs.dt2D,
        dtr: attrs.dtr,
        r1: attrs.r1,
        r2: attrs.r2,
      };
    }
  }

  setFormula(text: string): void {
    this.#formula = text;
    this.#hasFormula = true;
  }

  setValue(text: string): void {
    this.#valueText = text;
    this.#hasValue = true;
  }

  // Begin an `<is>`: clear the inline string and its runs so a rich value built from a previous cell's
  // runs keeps its own array.
  beginInlineString(): void {
    this.#inlineText = '';
    this.#runs.reset();
  }

  // Route a `<t>`'s text: to the open run when one is active, otherwise to the inline string when the
  // parser is inside an `<is>`. A run takes precedence — a run is also inside the inline string.
  appendText(text: string, inInlineString: boolean): void {
    if (!this.#runs.appendText(text) && inInlineString) this.#inlineText += text;
  }

  // Commit the gathered cell to the sheet with its already-resolved style (the caller applies the
  // cell → row → column inheritance order). A data-table cell surfaces its declaration; a shared-formula
  // master seeds the group before finalising, and a clone resolves to the master translated to its own
  // position. Everything else decodes as an ordinary `<c>` payload. An address-less cell is a no-op.
  finalize(
    sheet: Worksheet,
    sharedStrings: readonly SharedString[],
    style: XfStyle | undefined,
  ): void {
    if (this.#ref === '') return;
    if (this.#dataTable !== null) {
      const value: DataTableFormulaValue = {
        shareType: 'dataTable',
        ref: this.#dataTable.ref,
        ...(boolPresent(this.#dataTable.dt2D ?? '0') ? {dataTable2D: true} : {}),
        ...(boolPresent(this.#dataTable.dtr ?? '0') ? {dataTableRow: true} : {}),
        ...(this.#dataTable.r1 !== undefined ? {r1: this.#dataTable.r1} : {}),
        ...(this.#dataTable.r2 !== undefined ? {r2: this.#dataTable.r2} : {}),
        ...(this.#hasValue
          ? {result: decodeFormulaResult(this.#type, this.#valueText, style?.numFmt)}
          : {}),
      };
      const cell = sheet.getCell(this.#ref);
      applyXfToCell(cell, style);
      cell.value = value;
      return;
    }
    if (this.#hasFormula && this.#formulaShared && this.#formulaSi >= 0) {
      this.#masters.set(this.#formulaSi, {formula: this.#formula, col: this.#col, row: this.#row});
    } else if (this.#sharedClone && this.#formulaSi >= 0) {
      const master = this.#masters.get(this.#formulaSi);
      if (master !== undefined) {
        const translated = translateFormula(
          master.formula,
          this.#col - master.col,
          this.#row - master.row,
        );
        const value: SharedFormulaValue = {
          sharedFormula: encodeAddress(master.col, master.row),
          formula: unmangleFunctions(translated),
          // A clone's cached result honours the cell's date format the same way a plain formula's does.
          ...(this.#hasValue
            ? {result: decodeFormulaResult(this.#type, this.#valueText, style?.numFmt)}
            : {}),
        };
        const cell = sheet.getCell(this.#ref);
        applyXfToCell(cell, style);
        cell.value = value;
        return;
      }
    }
    const cell = sheet.getCell(this.#ref);
    applyXfToCell(cell, style);
    cell.value = this.decode(sharedStrings, style);
  }

  // Decode the gathered pieces into a plain cell value, resolving the shared pool and date formats but
  // NOT the shared-formula / data-table declarations {@link finalize} handles. This is what a data
  // read (the streaming reader) wants: the cell's own value, with a shared-formula clone surfacing its
  // cached result rather than a translated formula it will not evaluate.
  decode(sharedStrings: readonly SharedString[], style: XfStyle | undefined): CellValue {
    const raw: RawCell = {
      type: this.#type,
      hasFormula: this.#hasFormula,
      formula: this.#formula,
      hasValue: this.#hasValue,
      valueText: this.#valueText,
      inlineText: this.#inlineText,
      richTextRuns: this.#runs.runs,
    };
    return decodeCellContent(raw, sharedStrings, style?.numFmt);
  }
}

// Applies a resolved xf's non-value facets to a cell. Shared by the ordinary cell path and the
// shared-formula clone path, so a styled clone (fill/font/border/alignment/protection) keeps its
// look on read rather than surviving as value-only. The six cell-style facets go through the shared
// {@link applyCellStyle}; the xf-only links (`quotePrefix`, the named-style `xfId`) are applied here.
function applyXfToCell(cell: Cell, style: XfStyle | undefined): void {
  if (style === undefined) return;
  applyCellStyle(cell, style);
  if (style.quotePrefix !== undefined) cell.quotePrefix = style.quotePrefix;
  if (style.xfId !== undefined) cell.namedStyleId = style.xfId;
}
