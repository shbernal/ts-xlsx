// The resolved cell-format table both codecs read into: what an xf *is* once its id-indirection is
// flattened, which of the two number-format tables an id resolves against, and what applying an xf to
// a cell means.
//
// None of that is a property of how the style table is spelled — `xl/styles.xml` and `xl/styles.bin`
// state the same model in XML and BIFF12 records — so it is stated once here, above both codecs. The
// XML parsing lives in `../xlsx/read-styles.ts`, the record parsing in `../xlsb/read-styles.ts`, and
// each hands back the same {@link StyleTable}.

import {applyCellStyle, type Cell} from '../../core/cell.ts';
import {NAMED_STYLE_ID} from '../../core/internal.ts';
import type {CellStyle, Font, NamedCellStyle} from '../../core/style.ts';

/**
 * What an xf resolves to: the {@link CellStyle} facet tuple, plus the two flags an xf carries that
 * are not facets. Absent facets stay undefined, matching the contract that an unset facet is simply
 * not present on the reconstructed cell.
 *
 * It *derives* the facets rather than listing them, so a seventh facet added to `CellStyle` reaches
 * both readers the moment it joins. Re-declaring them here — the shape this replaced — meant a new
 * facet silently stopped at the model and never appeared in a file we read back.
 */
export interface XfStyle extends CellStyle {
  readonly quotePrefix?: boolean;
  /** The `xfId` link into the named-style layer (`cellStyleXfs`); absent for the Normal default (0). */
  readonly xfId?: number;
}

/** The parsed style table: the cell formats a cell/row/column `s` indexes, plus the named cell-style
 * layer a cell's `xfId` links into. Each cellXfs entry is already merged with its named style, so a
 * cell reading its `s` sees the effective facets; the `xfId` link is carried through for re-write. */
export interface StyleTable {
  readonly cellXfs: ReadonlyArray<XfStyle>;
  readonly namedStyles: ReadonlyArray<NamedCellStyle>;
  /**
   * Font id 0 — the workbook's declared default font, the face every cell naming no font renders in.
   * Surfaced separately from the fonts it was flattened onto because it is workbook-level state, not a
   * cell format: a re-write must emit *this* face as font 0 rather than an assumed Calibri, or every
   * empty cell changes face and every character-unit column width changes meaning. Absent when the
   * file declares no font table.
   */
  readonly defaultFont?: Font;
}

// ECMA-376 reserves numFmt ids below 164 for formats every consumer knows implicitly, so a
// foreign file may name one with no <numFmt> entry. This maps the standard ids to their
// codes; id 0 (General) and any unknown id resolve to no format. The writer never emits
// these — it always defines a custom id — but reading them keeps foreign files faithful.
const BUILTIN_NUMFMTS: ReadonlyMap<number, string> = new Map([
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'],
  [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
  // Ids 27..36 and 50..58 are reserved for locale-specific built-in East Asian date/time formats;
  // a file authored in a CJK locale styles date cells with them and, being built-ins, emits no
  // <numFmt>. The exact code is locale-defined — these are the representative Excel forms — but what
  // matters for reading is that each resolves to a non-empty date/time code so the serial reads as a
  // date rather than a bare number.
  [27, '[$-404]e/m/d'],
  [28, '[$-404]e"年"m"月"d"日"'],
  [29, '[$-404]e"年"m"月"d"日"'],
  [30, '[$-404]m/d/yy'],
  [31, '[$-404]yyyy"年"m"月"d"日"'],
  [32, '[$-404]h"時"mm"分"'],
  [33, '[$-404]h"時"mm"分"ss"秒"'],
  [34, '上午/下午h"時"mm"分"'],
  [35, '上午/下午h"時"mm"分"ss"秒"'],
  [36, '[$-404]e/m/d'],
  [50, '[$-404]e/m/d'],
  [51, '[$-404]e"年"m"月"d"日"'],
  [52, '[$-404]yyyy"年"m"月"'],
  [53, '[$-404]m"月"d"日"'],
  [54, '[$-404]e"年"m"月"d"日"'],
  [55, '上午/下午h"時"mm"分"'],
  [56, '上午/下午h"時"mm"分"ss"秒"'],
  [57, '[$-404]yyyy"年"m"月"'],
  [58, '[$-404]m"月"d"日"'],
]);

/**
 * The format code a number-format id denotes: the file's own `<numFmt>`/`BrtFmt` declaration if it
 * has one, else the built-in Excel defines for that id. Id 0 is General — the absence of a format —
 * and resolves to nothing so an ordinary cell carries no `numFmt`.
 */
export function numFmtCodeFor(id: number, custom: ReadonlyMap<number, string>): string | undefined {
  if (!Number.isInteger(id) || id === 0) return undefined;
  return custom.get(id) ?? BUILTIN_NUMFMTS.get(id);
}

/**
 * Apply a resolved xf's non-value facets to a cell — the six {@link CellStyle} facets through the
 * shared {@link applyCellStyle}, plus the two links that live on the xf itself rather than in the
 * facet tuple (`quotePrefix`, and the `xfId` pointer into the named-style layer).
 *
 * Shared by every path that commits a cell: the XML reader's ordinary and shared-formula-clone paths,
 * and the BIFF12 reader — so a styled cell keeps its look regardless of which serialisation it came
 * from, and the two cannot drift on what "applying a style" means.
 */
export function applyXfToCell(cell: Cell, style: XfStyle | undefined): void {
  if (style === undefined) return;
  applyCellStyle(cell, style);
  if (style.quotePrefix !== undefined) cell.quotePrefix = style.quotePrefix;
  if (style.xfId !== undefined) cell[NAMED_STYLE_ID] = style.xfId;
}
