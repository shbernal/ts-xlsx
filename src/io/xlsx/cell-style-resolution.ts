// Which `xf` record a cell's format resolves to, in one place for both readers.
//
// OOXML lets a cell inherit its format rather than state it: a `<c>` with no `s` takes its row's
// when the row is marked `customFormat`, and failing that its column's `<col style>`. That is the
// order Excel applies, and it is not decoration: `decodeCellContent` reads `numFmt` off the resolved
// style to decide whether a serial number is a *date*, so a cell whose format is inherited decodes
// to a different **type** depending on whether the resolution ran.
//
// `read-rows.ts` promises its decoded value is "identical to what `readXlsx` would produce for the
// same cell", and it was not: the buffered reader resolved all three levels and the streaming one
// read the cell's own `s` alone, so one date-formatted column turned `2023-03-15` into `45000` in
// exactly one of the two paths, silently. The rule now lives here and both drive it, which is the
// only arrangement under which that promise is checkable rather than merely stated.

import {boolStrict, numInteger} from '../../xml/xml-scan.ts';

/** The attributes this resolver reads, the subset both readers' scanners already hand it. */
interface StyleAttributes {
  readonly [key: string]: string | undefined;
}

/**
 * The cell → row → column style resolution for one worksheet, fed by the same elements in the same
 * order both readers already visit them: `<col>` before any cell references it, `<row>` before its
 * own cells.
 */
export class CellStyleResolver {
  // A column's `style` is the default for its cells that carry no style of their own: the index, not
  // the resolved record, because a `<col>` is parsed before the cells that inherit from it.
  readonly #columnStyle = new Map<number, number>();
  #rowStyle = -1;
  #rowCustomFormat = false;

  /** Record a `<col min max style>` span's default for the columns it covers. */
  noteColumnSpan(first: number, last: number, attrs: StyleAttributes): void {
    const styleIndex = numInteger(attrs.style, 0) ?? -1;
    if (styleIndex < 0) return;
    for (let index = first; index <= last; index++) this.#columnStyle.set(index, styleIndex);
  }

  /** Open a `<row>`, taking the format it declares for its own bare cells. */
  openRow(attrs: StyleAttributes): void {
    this.#rowStyle = numInteger(attrs.s, 0) ?? -1;
    this.#rowCustomFormat = boolStrict(attrs.customFormat);
  }

  /** Close the row, so a cell outside one never inherits the last row's format. */
  closeRow(): void {
    this.#rowStyle = -1;
    this.#rowCustomFormat = false;
  }

  /**
   * The `xfStyles` index a cell's format resolves to, or `-1` when nothing declares one. A row's
   * format applies only when the row is marked `customFormat`; without that flag the `s` on a `<row>`
   * describes the row itself and is not inherited by its cells.
   */
  indexFor(col: number, cellStyleIndex: number): number {
    if (cellStyleIndex >= 0) return cellStyleIndex;
    if (this.#rowCustomFormat && this.#rowStyle >= 0) return this.#rowStyle;
    return this.#columnStyle.get(col) ?? -1;
  }
}
