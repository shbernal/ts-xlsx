// Rendering one row and its cells: the `<row>` element, the `<c>` elements inside it, and everything
// a cell's own value and format contribute to them.
//
// Split from `worksheet-xml.ts`, which stays the sheet orchestrator: the two halves shared nothing but
// the style registry, and this is precisely the surface the streaming writer imports. That writer
// serialises a row the moment it is committed, long before the sheet around it exists, so "what does
// one row look like" is a question with an answer independent of the sheet, and a module boundary is
// the honest way to say so.

import {type Cell, cellHasOwnStyle} from '../../core/cell.ts';
import {DEFAULT_DATE_NUMFMT, dateToSerial} from '../../core/date.ts';
import {mangleFormula} from '../../core/formula.ts';
import {NAMED_STYLE_ID} from '../../core/internal.ts';
import {CELL_STYLE_FACETS, type CellStyle, type Fill} from '../../core/style.ts';
import {
  type FormulaResult,
  detectValueType,
  isDataTableFormulaValue,
  isErrorValue,
  isFormulaValue,
  isHyperlinkValue,
  isRichTextValue,
  isSharedFormulaValue,
} from '../../core/value.ts';
import type {ColumnProperties, RowProperties, Worksheet} from '../../core/worksheet.ts';
import {InternalError} from '../../errors.ts';
import {
  escapeAttr,
  escapeSpreadsheetText,
  escapeText,
  numberText,
  textAttr,
  textElement,
} from '../../xml/xml.ts';
import type {XfStyle} from '../style/xf-style.ts';
import type {CommentCell} from './comments.ts';
import type {CollectedHyperlink} from './hyperlinks.ts';
import {richTextRunsXml} from './rich-text.ts';
import type {SharedFormulaRole} from './shared-formulas.ts';
import type {SharedStringTable} from './shared-strings.ts';
import type {StyleRegistry} from './styles.ts';

/**
 * The used-cell extent of a sheet: the top-left/bottom-right grid bounds that fold into the
 * `<dimension>`. Rows carrying only formatting (a row height, an outline level) do not extend the
 * used range, matching how Excel records `<dimension>`, so {@link add} ignores them. A fresh extent
 * holds the `Infinity`/`-Infinity` sentinels; {@link isEmpty} reports that no used cell has been seen.
 */
export class Extent {
  top = Infinity;
  left = Infinity;
  bottom = -Infinity;
  right = -Infinity;

  // Seed from a prior extent (the rows a streaming writer already flushed and evicted) so the buffered
  // pass folds its live rows onto the same bounds; unseeded, it starts empty.
  constructor(seed?: Extent) {
    if (seed) {
      this.top = seed.top;
      this.left = seed.left;
      this.bottom = seed.bottom;
      this.right = seed.right;
    }
  }

  /** Whether no used cell has been folded in yet, in which case the dimension is the lone cell `A1`. */
  get isEmpty(): boolean {
    return this.bottom === -Infinity;
  }

  /** Fold a rendered row's used-column span into the extent. `minCol` is `Infinity` when the row
   * carried no cells (only formatting), which extends nothing. */
  add(row: number, minCol: number, maxCol: number): void {
    if (minCol === Infinity) return;
    if (row < this.top) this.top = row;
    if (row > this.bottom) this.bottom = row;
    if (minCol < this.left) this.left = minCol;
    if (maxCol > this.right) this.right = maxCol;
  }
}

/**
 * A worksheet's eagerly-serialised rows: each row's `<row>` XML tagged with its number (so it merges
 * into ascending order with the sheet's remaining live rows, whatever order it was committed in), plus
 * the used-cell {@link Extent} they span. The buffered pass folds that extent into the sheet's dimension.
 */
export interface FlushedSheet {
  readonly rows: ReadonlyArray<{readonly number: number; readonly xml: string}>;
  readonly extent: Extent;
  /**
   * Each flushed row's outline level and hidden flag. Carried across the eviction because both feed
   * whole-sheet derivations made long after the row is gone from the model: `<sheetFormatPr
   * outlineLevelRow>` is the deepest level on the sheet, and `collapsed="1"` rides a summary row only
   * when its whole detail group is hidden, which is a question about *other* rows.
   *
   * Carried as the two inputs rather than the derived flag, because the derivation is a look-ahead a
   * row being finalised cannot do, while its inputs are per-row and known exactly when the row flushes.
   */
  readonly rowOutline: ReadonlyMap<
    number,
    {readonly outlineLevel: number; readonly hidden: boolean}
  >;
  /**
   * The hyperlinks and notes the flushed rows carried, on the same terms and for the same reason.
   * Both are serialised outside the `<row>`: a hyperlink into the sheet's `<hyperlinks>` element plus
   * an external relationship, a note into the comments and VML parts. The buffered pass gathers them
   * by walking the sheet's rows at commit time, which finds nothing on a row whose cells have already
   * been evicted, so a streamed row's link kept its visible label and silently lost its destination.
   */
  readonly hyperlinks: readonly CollectedHyperlink[];
  readonly notes: readonly CommentCell[];
}

/**
 * A column's style facets are defaults its cells inherit unless they override them; the writer
 * composes each cell's full style up front (cell over row over column, per facet) so a cell that
 * overrides one facet still carries the column's others, rather than silently dropping them. Frozen
 * once by the streaming writer at its first flush so every eagerly-rendered row sees the same defaults.
 */
export function buildColumnDefaults(sheet: Worksheet): Map<number, ColumnProperties> {
  const columnDefaults = new Map<number, ColumnProperties>();
  // `columns()` yields only columns that carry a format record, so the fallback is unreachable.
  // it is here because the handle's `properties` is honestly optional, not because a defined
  // column can lack one.
  for (const {index, properties} of sheet.columns()) columnDefaults.set(index, properties ?? {});
  return columnDefaults;
}

/** The whole-sheet context a single row needs to serialise: the column defaults it inherits, the
 * style/string tables it interns into, the shared-formula roles its cells play, and the collapsed
 * outline summaries whose toggle it must stamp. The streaming writer supplies empty shared-formula
 * and collapsed-summary sets, since those are whole-sheet derivations a flushed row cannot join. */
export interface RowRenderContext {
  readonly columnDefaults: ReadonlyMap<number, ColumnProperties>;
  readonly styles: StyleRegistry;
  readonly sharedStrings: SharedStringTable | null;
  readonly sharedRoles: ReadonlyMap<string, SharedFormulaRole>;
  readonly collapsedSummaries: ReadonlySet<number>;
}

/**
 * Serialise one row to its `<row>` element, or '' when the row has neither data nor its own
 * formatting. Returns the used-column bounds (`Infinity`/`-Infinity` when nothing was rendered) so a
 * caller can fold them into the sheet dimension. Shared by the buffered sheet pass and the streaming
 * writer's eager flush, so both emit byte-identical rows.
 */
export function renderRow(
  entry: {
    readonly number: number;
    readonly cells: readonly Cell[];
    readonly properties: RowProperties | undefined;
  },
  ctx: RowRenderContext,
): {xml: string; minCol: number; maxCol: number} {
  const {number, cells, properties} = entry;
  // A cell earns a <c> element if it holds a value OR carries its own style: a formatted-but-empty
  // cell (a fill/border on a null value) is a real cell to Excel, and dropping it would lose the
  // formatting. A cell with neither is inherited from its row/column and needs no element of its own.
  const rendered = cells.filter((cell) => cell.value !== null || cellHasOwnStyle(cell));
  const attrs = rowAttrs(properties, ctx.styles, ctx.collapsedSummaries.has(number));
  // A row with neither data nor its own formatting has nothing to serialise.
  if (rendered.length === 0 && attrs === '') return {xml: '', minCol: Infinity, maxCol: -Infinity};
  const rowFill = properties?.fill;
  const cellsXml = rendered
    .map((cell) => {
      const style = ctx.styles.styleId(
        composeCellStyle(cell, rowFill, ctx.columnDefaults.get(cell.col)),
      );
      return cellXml(cell, style, ctx.sharedRoles.get(cell.address), ctx.sharedStrings);
    })
    .join('');
  let minCol = Infinity;
  let maxCol = -Infinity;
  for (const cell of rendered) {
    if (cell.col < minCol) minCol = cell.col;
    if (cell.col > maxCol) maxCol = cell.col;
  }
  return {xml: `<row r="${number}"${attrs}>${cellsXml}</row>`, minCol, maxCol};
}

// Compose a cell's full style by resolving each facet cell-over-row-over-column, so a cell that
// overrides one facet still carries the row's fill and the column's other facets rather than silently
// dropping them: the per-facet precedence Excel applies.
//
// Cell-over-column is the default, taken from the facet registry, so a seventh facet added to
// `CellStyle` composes correctly on the day it joins. The two facets that resolve against more than
// that overwrite themselves afterwards rather than being exempted from the loop: the loop stays the
// exhaustive half, and a special case has to be written down to exist. Quote-prefix and the
// named-style link are cell-only, with no row/column default to inherit, and are deliberately not
// facets.
function composeCellStyle(
  cell: Cell,
  rowFill: Fill | undefined,
  colDef: ColumnProperties | undefined,
): XfStyle {
  const style: {-readonly [K in keyof XfStyle]?: XfStyle[K]} = {};
  for (const facet of CELL_STYLE_FACETS) inheritFacet(style, facet, cell, colDef);
  // A row carries a fill and nothing else, so it sits between the cell and the column on this one
  // facet alone.
  style.fill = cell.fill ?? rowFill ?? colDef?.fill;
  // A bare Date carries no format of its own, so it renders as a raw serial and reads back as a
  // number unless we apply a date format. An explicit cell/column format wins.
  style.numFmt = cell.numFmt ?? colDef?.numFmt ?? dateDefaultNumFmt(cell.value);
  if (cell.quotePrefix !== undefined) style.quotePrefix = cell.quotePrefix;
  // Preserved so a round-trip keeps the cell tied to its named style rather than flattening it into
  // a purely-direct format.
  const xfId = cell[NAMED_STYLE_ID];
  if (xfId !== undefined) style.xfId = xfId;
  return style;
}

// One facet at a time, so `target[key] = a[key] ?? b[key]` typechecks without a cast: the
// correlated-key access TS cannot verify when the key is the whole union. The same shape, and the
// same reason, as `copyFacet` in `core/style.ts`.
function inheritFacet<K extends keyof CellStyle>(
  target: CellStyle,
  key: K,
  cell: Readonly<CellStyle>,
  colDef: Readonly<CellStyle> | undefined,
): void {
  target[key] = cell[key] ?? colDef?.[key];
}

function rowAttrs(
  properties: RowProperties | undefined,
  styles: StyleRegistry,
  collapsedSummary: boolean,
): string {
  if (properties === undefined) return collapsedSummary ? ' collapsed="1"' : '';
  let attrs = '';
  if (properties.height !== undefined)
    attrs += ` ht="${numberText(properties.height)}" customHeight="1"`;
  if (properties.hidden) attrs += ' hidden="1"';
  attrs += outlineAttr('outlineLevel', properties.outlineLevel);
  // The collapse toggle is set explicitly by the author, or derived onto a summary row whose whole
  // detail group is hidden (the `collapsedSummaries` set {@link scanRowOutline} builds). It rides the
  // summary row, never the detail rows.
  if (properties.collapsed || collapsedSummary) attrs += ' collapsed="1"';
  // A row-level fill is a default format for the row's cells; customFormat="1" is what makes
  // Excel honour the row's `s`, and a cell without its own `s` then inherits it.
  const style = styles.styleId({fill: properties.fill});
  if (style !== 0) attrs += ` s="${style}" customFormat="1"`;
  return attrs;
}

// A valid Date, whether the cell's own value or a formula's cached result, with no format of its
// own gets the default date format so it renders and reads back as a date rather than a bare serial.
// An Invalid Date and every non-date value contribute nothing here.
function dateDefaultNumFmt(value: Cell['value']): string | undefined {
  const date =
    value instanceof Date
      ? value
      : (isFormulaValue(value) || isSharedFormulaValue(value)) && value.result instanceof Date
        ? value.result
        : undefined;
  return date !== undefined && !Number.isNaN(date.getTime()) ? DEFAULT_DATE_NUMFMT : undefined;
}

function cellXml(
  cell: Cell,
  style: number,
  shared: SharedFormulaRole | undefined,
  sharedStrings: SharedStringTable | null,
): string {
  const ref = cell.address;
  const value = cell.value;
  const s = style !== 0 ? ` s="${style}"` : '';

  const formula = cellFormulaXml(ref, s, value, shared);
  if (formula !== undefined) return formula;

  if (value instanceof Date) {
    // An Invalid Date (new Date(NaN)) has no serial; keep the cell (and its style) but emit no
    // value rather than throwing, so one bad date never takes down the whole sheet's export.
    if (Number.isNaN(value.getTime())) return `<c r="${ref}"${s}/>`;
    return `<c r="${ref}"${s}><v>${numberText(dateToSerial(value))}</v></c>`;
  }
  if (typeof value === 'number') {
    // A non-finite number (NaN, ±Infinity) has no OOXML representation; keep the cell and its style
    // but emit no value rather than a bare "NaN"/"Infinity" token: the same graceful degradation an
    // Invalid Date gets, so one bad value never corrupts the sheet or takes down the whole export.
    if (!Number.isFinite(value)) return `<c r="${ref}"${s}/>`;
    return `<c r="${ref}"${s}><v>${numberText(value)}</v></c>`;
  }
  if (typeof value === 'boolean') {
    return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  if (typeof value === 'string') {
    // With shared strings on, the cell holds only the pool index (`t="s"`); otherwise the text
    // lives inline in the cell. Both decode to the same string on read.
    if (sharedStrings !== null) {
      return `<c r="${ref}"${s} t="s"><v>${sharedStrings.intern(value)}</v></c>`;
    }
    return `<c r="${ref}"${s} t="inlineStr"><is>${textElement(value)}</is></c>`;
  }
  if (isRichTextValue(value)) {
    // With shared strings on, rich text is pooled as a rich `<si>` (the cell holds only its index);
    // otherwise the runs live inline. Both decode back to the same runs on read.
    if (sharedStrings !== null) {
      return `<c r="${ref}"${s} t="s"><v>${sharedStrings.intern(value)}</v></c>`;
    }
    return `<c r="${ref}"${s} t="inlineStr"><is>${richTextRunsXml(value.richText)}</is></c>`;
  }
  if (isHyperlinkValue(value)) {
    // The cell holds only the visible label; the link itself rides in the sheet's <hyperlinks>.
    // The label is either a plain string or rich text, serialised the same way a cell value of
    // that kind would be.
    const label =
      typeof value.text === 'string'
        ? textElement(value.text)
        : richTextRunsXml(value.text.richText);
    return `<c r="${ref}"${s} t="inlineStr"><is>${label}</is></c>`;
  }
  if (isErrorValue(value)) {
    // An error literal serialises under t="e" with its code as the value. The codes are a closed
    // set of canonical spellings (see ERROR_CODES) with no XML-special characters, so no escaping.
    return `<c r="${ref}"${s} t="e"><v>${value.error}</v></c>`;
  }
  // A null value only reaches here for a formatted-but-empty cell (the row loop keeps it for its
  // style); emit the styled cell with no <v>, exactly how Excel stores a formatted blank.
  if (value === null) return `<c r="${ref}"${s}/>`;
  // Every ValueType kind is served by an arm above (a formula routes through its own writer), so
  // this is unreachable. It exists because the union is not exhaustively narrowed here.
  throw new InternalError(
    `writing a ${detectValueType(value)} cell value has no arm: every CellValue kind is handled above`,
  );
}

// Serialise a formula cell (a shared-formula master or clone, a What-If data table, or a plain
// formula) into its `<c>` element, or return undefined when the value is not a formula so `cellXml`
// falls through to its value dispatch.
function cellFormulaXml(
  ref: string,
  s: string,
  value: Cell['value'],
  shared: SharedFormulaRole | undefined,
): string | undefined {
  // A shared-formula master seeds the group with its formula text under `t="shared" ref si`; a clone
  // carries no text of its own, only a back-reference to the master's `si`. Its cached result still
  // travels with the cell.
  if (shared !== undefined) {
    if (shared.ref !== undefined && isFormulaValue(value)) {
      const f = `<f t="shared" ref="${shared.ref}" si="${shared.si}">${escapeText(mangleFormula(value.formula))}</f>`;
      return formulaBodyXml(ref, s, f, value.result);
    }
    const result = isSharedFormulaValue(value) ? value.result : undefined;
    return formulaBodyXml(ref, s, `<f t="shared" si="${shared.si}"/>`, result);
  }
  if (isDataTableFormulaValue(value)) {
    // A data-table formula carries no expression text, only its declaration attributes, which we
    // re-emit verbatim so a read-modify-write cycle preserves the What-If kind the library never
    // evaluates. The cached result travels as any formula result does.
    const attrs =
      `ref="${escapeAttr(value.ref)}"` +
      ` dt2D="${value.dataTable2D ? 1 : 0}"` +
      ` dtr="${value.dataTableRow ? 1 : 0}"` +
      textAttr('r1', value.r1) +
      textAttr('r2', value.r2);
    return formulaBodyXml(ref, s, `<f t="dataTable" ${attrs}/>`, value.result);
  }
  if (isFormulaValue(value)) {
    return formulaBodyXml(
      ref,
      s,
      `<f>${escapeText(mangleFormula(value.formula))}</f>`,
      value.result,
    );
  }
  return undefined;
}

// Wrap a prepared `<f>` element (a plain formula, or a shared master/slave `<f>`) with the cell
// element and its cached result, typing the cell by the result's kind exactly as a bare value of that
// kind would be.
function formulaBodyXml(
  ref: string,
  s: string,
  f: string,
  result: FormulaResult | undefined,
): string {
  // A non-finite cached result (a `1/0` that reached the model as Infinity/NaN) has no OOXML
  // representation; keep the formula but cache no value rather than emit a bare "NaN": the same
  // graceful degradation a bare non-finite cell and an Invalid Date result get.
  if (result === undefined || (typeof result === 'number' && !Number.isFinite(result))) {
    return `<c r="${ref}"${s}>${f}</c>`;
  }
  if (typeof result === 'number') {
    return `<c r="${ref}"${s}>${f}<v>${numberText(result)}</v></c>`;
  }
  if (typeof result === 'boolean') {
    return `<c r="${ref}"${s} t="b">${f}<v>${result ? 1 : 0}</v></c>`;
  }
  if (typeof result === 'string') {
    // The cached result of a string formula is a cell value, not structure, so it carries the
    // `_xHHHH_` escape a `<t>` does, and Excel decodes it here too (verified over COM: a `<v>` of
    // `_x0041_` under t="str" reads back as "A" with calculation held manual).
    return `<c r="${ref}"${s} t="str">${f}<v>${escapeSpreadsheetText(result)}</v></c>`;
  }
  if (isErrorValue(result)) {
    // A formula that evaluated to an error caches its code under t="e", exactly as a bare error
    // cell does: the reader's decodeResult mirrors decodeValue for this case.
    return `<c r="${ref}"${s} t="e">${f}<v>${result.error}</v></c>`;
  }
  if (result instanceof Date) {
    // A date-valued result caches its serial exactly as a bare date cell stores its value; the
    // cell's date number format (applied when its style is composed) is what makes both read back as
    // a Date. An Invalid Date has no serial, so cache no result rather than emit NaN.
    if (Number.isNaN(result.getTime())) return `<c r="${ref}"${s}>${f}</c>`;
    return `<c r="${ref}"${s}>${f}<v>${numberText(dateToSerial(result))}</v></c>`;
  }
  // Every FormulaResult kind is handled above; this guards a value that reached here past the model.
  throw new InternalError(
    'writing a non-primitive formula result has no arm: every FormulaResult kind is handled above',
  );
}

// An outline depth is written only above the default of zero. That zero test answers whether the
// level is worth recording, not whether it can be recorded at all, and keeping the two apart is what
// makes an unwritable level loud: `NaN > 0` is false, so a gate on its own would drop it silently
// while `Infinity` sailed through into an `xsd:unsignedInt` attribute.
export function outlineAttr(name: string, level: number | undefined): string {
  if (level === undefined) return '';
  const text = numberText(level);
  return level > 0 ? ` ${name}="${text}"` : '';
}
