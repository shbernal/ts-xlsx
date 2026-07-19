// The worksheet-body reader: one `xl/worksheets/sheetN.xml` in, its cells/columns/rows/print-setup
// applied onto a {@link Worksheet}. It is a single streaming pass driving a handful of accumulators
// (the cell being read, shared-formula masters, an autofilter draft, the current page-break axis) so
// each element commits its state as it closes. Style indices resolve through the parsed style table.

import {decodeAddress, decodeRange, encodeAddress} from '../../core/address.ts';
import {
  type CustomFilterPredicate,
  type FilterColumn,
  type FilterCriteria,
  isCustomFilterOperator,
} from '../../core/autofilter.ts';
import type {Cell} from '../../core/cell.ts';
import {translateFormula, unmangleFunctions} from '../../core/formula.ts';
import type {PageBreak, PageMargins, PageSetup, PrintOptions} from '../../core/page-setup.ts';
import {
  SHEET_PROTECTION_FLAGS,
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionFlags,
} from '../../core/protection.ts';
import type {DataTableFormulaValue, SharedFormulaValue} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {
  decodeCellContent,
  decodeFormulaResult,
  type RawCell,
  type SharedString,
} from './cell-value.ts';
import type {XfStyle} from './read-styles.ts';
import {RunAccumulator} from './rich-runs.ts';
import {parseColor} from './styles.ts';
import {
  boolPresent,
  boolStrict,
  boolTristate,
  localName,
  parseXml,
  type XmlAttributes,
} from './xml-read.ts';

const MARGIN_SIDES = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const;

// Worksheet elements that commit on their close: a formatted-but-empty `<c/>` and a criteria-free
// self-closing `<autoFilter/>` are expanded to open+close so each finalises once in onClose. The
// text-bearing `<f/>`/`<v/>`/`<t/>` are deliberately excluded — an empty one must not commit.
const WORKSHEET_EMPTY_CLOSES: ReadonlySet<string> = new Set(['c', 'autoFilter']);

// Fold a filter column's accumulated `<filters>` or `<customFilters>` state into one criteria value,
// or null when it carried nothing filterable. A `<filters>` block with no values and no blank flag,
// or a `<customFilters>` with no predicates, is a no-op that would round-trip as noise, so it drops.
function pendingFilterCriteria(
  values: string[] | null,
  blank: boolean,
  predicates: CustomFilterPredicate[] | null,
  and: boolean,
): FilterCriteria | null {
  if (values !== null && (values.length > 0 || blank)) {
    return {kind: 'values', values, blank};
  }
  if (predicates !== null && predicates.length > 0) {
    return {kind: 'custom', and, predicates: predicates.slice(0, 2)};
  }
  return null;
}

// Read a <sheetProtection> element back into a SheetProtection — the deserialization mirror of the
// writer. `sheet="0"` (or "false") means the element records an *un*protected sheet, so nothing is
// restored. Each flag attribute is the INVERSE of the author's allow-flag ("1" forbids, "0" permits),
// and only attributes actually present are carried, so an omitted (default-valued) flag stays absent —
// exactly what the writer emitted. A password credential is preserved verbatim in its agile form
// (algorithm, hash, salt, spin count); there is no plaintext password to recover, so it is not re-hashed.
function parseSheetProtection(attrs: {readonly [k: string]: string}): SheetProtection | undefined {
  if (attrs.sheet === '0' || attrs.sheet === 'false') return undefined;
  const flags: {-readonly [K in keyof SheetProtectionFlags]?: boolean} = {};
  for (const {key} of SHEET_PROTECTION_FLAGS) {
    const raw = attrs[key];
    if (raw !== undefined) flags[key] = !boolStrict(raw);
  }
  const {algorithmName, hashValue, saltValue, spinCount} = attrs;
  if (
    algorithmName !== undefined &&
    hashValue !== undefined &&
    saltValue !== undefined &&
    spinCount !== undefined
  ) {
    const spin = Number(spinCount);
    if (Number.isFinite(spin)) {
      const credential: SheetProtectionCredential = {
        algorithmName,
        hashValue,
        saltValue,
        spinCount: spin,
      };
      return {flags, credential};
    }
  }
  return {flags};
}

// Autofilter accumulation. The sheet `<autoFilter ref>` seeds a draft; each `<filterColumn colId>`
// opens a column whose criteria (`<filters>` values or `<customFilters>` predicates) stream into the
// draft until `</filterColumn>`, and `</autoFilter>` commits the whole thing to the sheet.
class AutoFilterAccumulator {
  #ref: string | null = null;
  #columns: FilterColumn[] = [];
  #colId = -1;
  #values: string[] | null = null;
  #blank = false;
  #predicates: CustomFilterPredicate[] | null = null;
  #and = false;

  // Seed a draft from the range. (A table's own `<autoFilter>` also matches, but table sheets route
  // through parseTable, so this only ever sees the sheet-level one.)
  begin(attrs: XmlAttributes): void {
    this.#ref = attrs.ref !== undefined && attrs.ref !== '' ? attrs.ref : null;
    this.#columns = [];
  }

  // Open a criteria block for one column, offset `colId` from the range's left edge. Reset the
  // per-column accumulators; whichever child (`<filters>`/`<customFilters>`) opens fills one.
  beginColumn(attrs: XmlAttributes): void {
    this.#colId = attrs.colId !== undefined ? Number(attrs.colId) : -1;
    this.#values = null;
    this.#blank = false;
    this.#predicates = null;
    this.#and = false;
  }

  beginValues(attrs: XmlAttributes): void {
    this.#values = [];
    this.#blank = boolPresent(attrs.blank) && attrs.blank !== undefined;
  }

  addValue(attrs: XmlAttributes): void {
    if (this.#values !== null && attrs.val !== undefined) this.#values.push(attrs.val);
  }

  beginCustom(attrs: XmlAttributes): void {
    this.#predicates = [];
    this.#and = attrs.and !== undefined && boolPresent(attrs.and);
  }

  // The operator attribute defaults to `equal` when absent (per CT_CustomFilter); an operand is
  // likewise optional. An unrecognised operator drops the predicate rather than guessing.
  addCustom(attrs: XmlAttributes): void {
    if (this.#predicates === null) return;
    const operator = attrs.operator ?? 'equal';
    if (isCustomFilterOperator(operator)) {
      this.#predicates.push({operator, val: attrs.val ?? ''});
    }
  }

  // Assemble this column's criteria from whichever accumulator filled. A column whose colId is
  // negative, or whose criteria are empty (no values, no blank, no predicates), carries nothing
  // filterable and is dropped so a re-write stays clean — load-repair, not authoring.
  endColumn(): void {
    const criteria = pendingFilterCriteria(this.#values, this.#blank, this.#predicates, this.#and);
    if (this.#colId >= 0 && criteria !== null) {
      this.#columns.push({colId: this.#colId, criteria});
    }
  }

  // Commit the accumulated autofilter to the sheet. Runs on `</autoFilter>` — including the synthesized
  // close of a criteria-free self-closing `<autoFilter/>`. Columns whose colId falls outside the range
  // are dropped here so the strict setter never trips on hostile input.
  commit(sheet: Worksheet): void {
    if (this.#ref === null) return;
    try {
      const {left, right} = decodeRange(this.#ref);
      const width = left !== undefined && right !== undefined ? right - left + 1 : 0;
      sheet.autoFilter = {ref: this.#ref, columns: this.#columns.filter((c) => c.colId < width)};
    } catch {
      // unbounded or malformed autofilter range in the source file — ignore it
    }
    this.#ref = null;
    this.#columns = [];
  }
}

// Page-break accumulation. `<brk>` elements appear under both `<rowBreaks>` and `<colBreaks>`; a break
// container's open points the accumulator at that axis's list (null outside any container), so a
// `<brk>` lands on the right axis, and the matching close clears it. A self-closing
// `<rowBreaks/>`/`<colBreaks/>` fires no close, so a new open simply reassigns the target.
class PageBreakAccumulator {
  #target: PageBreak[] | null = null;

  begin(target: PageBreak[]): void {
    this.#target = target;
  }

  end(): void {
    this.#target = null;
  }

  // `id` is the row/column the layout splits before; a non-positive or non-integer id is hostile input
  // and dropped rather than trusted. A `<brk>` outside any break container has no axis and is ignored.
  add(attrs: XmlAttributes): void {
    if (this.#target === null) return;
    const id = Number(attrs.id);
    if (!Number.isInteger(id) || id < 1) return;
    const brk: {id: number; max?: number; man?: boolean} = {id};
    const max = Number(attrs.max);
    if (Number.isInteger(max) && max >= 0) brk.max = max;
    if (boolStrict(attrs.man)) brk.man = true;
    this.#target.push(brk);
  }
}

export function parseWorksheet(
  xml: string,
  sheet: Worksheet,
  sharedStrings: readonly SharedString[],
  xfStyles: ReadonlyArray<XfStyle>,
): void {
  let cellRef = '';
  let cellType = '';
  let cellStyle = -1;
  let cellCol = -1;
  let cellRow = -1;
  let formula = '';
  // Shared-formula bookkeeping. A master `<f t="shared" ref si>TEXT</f>` seeds the group; every clone
  // `<f t="shared" si/>` in the sheet references it by `si` and carries no text of its own. Masters
  // always precede their clones (Excel keeps the master top-left), so a clone resolves against a map
  // filled as the sheet streams: the master's formula translated to the clone's position.
  const sharedMasters = new Map<number, {formula: string; col: number; row: number}>();
  let formulaShared = false;
  let formulaSi = -1;
  let sharedClone = false;
  // The declaration attributes of a `<f t="dataTable">`, held from the `<f>` open until the cell
  // finalises. Null for every other cell.
  let formulaDataTable: {
    ref: string;
    dt2D: string | undefined;
    dtr: string | undefined;
    r1: string | undefined;
    r2: string | undefined;
  } | null = null;
  let valueText = '';
  let inlineText = '';
  let hasFormula = false;
  let hasValue = false;
  let inInlineString = false;
  let capture = false;
  let text = '';
  // Rich-text run accumulation while inside an `<is>` built from `<r>` runs. A `<t>` inside a run
  // appends to the run's text; a bare `<t>` directly in the `<is>` appends to `inlineText`, keeping a
  // plain inline string on its existing path.
  const runs = new RunAccumulator();
  // A row with customFormat="1" supplies a default style for its cells that carry no `s`.
  let rowStyle = -1;
  let rowCustomFormat = false;
  const autoFilter = new AutoFilterAccumulator();
  const pageBreaks = new PageBreakAccumulator();
  // A column's `style` is the default for its cells that carry no style of their own; this
  // maps a column index to that style index so a bare cell can inherit it (as Excel does,
  // without stamping every cell). Columns are parsed before any cell references them.
  const columnStyle = new Map<number, number>();

  // Commit the cell held in the parser state to the sheet, resolving its style from its own `s`,
  // then its row's (when customFormat), then its column's default — the order Excel applies. Runs on
  // `</c>` close, including the synthesized close of a self-closing `<c/>` formatted-but-empty cell.
  const finalizeCellFromState = (): void => {
    if (cellRef === '') return;
    const styleIndex =
      cellStyle >= 0
        ? cellStyle
        : rowCustomFormat && rowStyle >= 0
          ? rowStyle
          : (columnStyle.get(cellCol) ?? -1);
    const style = styleIndex >= 0 ? xfStyles[styleIndex] : xfStyles[0];
    // A data-table formula is preserved by declaration, not evaluated: surface its kind, range, input
    // cells, and cached result so a read-modify-write cycle re-emits it rather than dropping it.
    if (formulaDataTable !== null) {
      const value: DataTableFormulaValue = {
        shareType: 'dataTable',
        ref: formulaDataTable.ref,
        ...(boolPresent(formulaDataTable.dt2D ?? '0') ? {dataTable2D: true} : {}),
        ...(boolPresent(formulaDataTable.dtr ?? '0') ? {dataTableRow: true} : {}),
        ...(formulaDataTable.r1 !== undefined ? {r1: formulaDataTable.r1} : {}),
        ...(formulaDataTable.r2 !== undefined ? {r2: formulaDataTable.r2} : {}),
        ...(hasValue ? {result: decodeFormulaResult(cellType, valueText, style?.numFmt)} : {}),
      };
      const dtCell = sheet.getCell(cellRef);
      applyCellStyle(dtCell, style);
      dtCell.value = value;
      return;
    }
    // A shared-formula master seeds the group for its clones before it is finalised as an ordinary
    // formula cell; a clone resolves to the master's formula translated to its own position and is
    // committed here directly, since its value is not a plain `<c>` payload decodeCellContent knows.
    if (hasFormula && formulaShared && formulaSi >= 0) {
      sharedMasters.set(formulaSi, {formula, col: cellCol, row: cellRow});
    } else if (sharedClone && formulaSi >= 0) {
      const master = sharedMasters.get(formulaSi);
      if (master !== undefined) {
        const translated = translateFormula(
          master.formula,
          cellCol - master.col,
          cellRow - master.row,
        );
        const value: SharedFormulaValue = {
          sharedFormula: encodeAddress(master.col, master.row),
          formula: unmangleFunctions(translated),
          // A clone's cached result honours the cell's date format the same way a plain formula's does.
          ...(hasValue ? {result: decodeFormulaResult(cellType, valueText, style?.numFmt)} : {}),
        };
        const cloneCell = sheet.getCell(cellRef);
        applyCellStyle(cloneCell, style);
        cloneCell.value = value;
        return;
      }
    }
    finalizeCell(
      sheet,
      cellRef,
      {
        type: cellType,
        hasFormula,
        formula,
        hasValue,
        valueText,
        inlineText,
        richTextRuns: runs.runs,
      },
      sharedStrings,
      style,
    );
  };

  parseXml(
    xml,
    {
      onOpen(name, attrs, selfClosing) {
        const local = localName(name);
        text = '';
        capture = false;
        switch (local) {
          case 'col':
            applyColumn(sheet, attrs, xfStyles, columnStyle);
            break;
          case 'row':
            applyRow(sheet, attrs);
            rowStyle = attrs.s !== undefined ? Number(attrs.s) : -1;
            rowCustomFormat = boolStrict(attrs.customFormat);
            break;
          case 'c':
            cellRef = attrs.r ?? '';
            cellType = attrs.t ?? '';
            cellStyle = attrs.s !== undefined ? Number(attrs.s) : -1;
            cellCol = cellRef === '' ? -1 : (decodeAddress(cellRef).col ?? -1);
            cellRow = cellRef === '' ? -1 : (decodeAddress(cellRef).row ?? -1);
            formula = '';
            valueText = '';
            inlineText = '';
            runs.reset();
            hasFormula = false;
            hasValue = false;
            formulaShared = false;
            formulaSi = -1;
            sharedClone = false;
            formulaDataTable = null;
            break;
          case 'is':
            inInlineString = true;
            inlineText = '';
            runs.reset();
            break;
          case 'r':
            // A run inside a rich inline string. Its `<rPr>` (if any) and `<t>` follow.
            if (inInlineString) runs.beginRun();
            break;
          case 'rPr':
            // The run's formatting bundle; its self-closing children stream into the default branch.
            runs.beginProperties();
            break;
          case 'f':
            capture = true;
            formulaShared = attrs.t === 'shared';
            formulaSi = attrs.si !== undefined ? Number(attrs.si) : -1;
            // A self-closing `<f t="shared" si/>` is a clone: it fires no close event and carries no
            // text, so mark it here to resolve against its master when the cell finalises.
            if (selfClosing && formulaShared) sharedClone = true;
            // A `<f t="dataTable">` carries only declaration attributes; hold them for finalisation so
            // the data-table kind is preserved rather than read as an empty formula.
            if (attrs.t === 'dataTable' && attrs.ref !== undefined) {
              formulaDataTable = {
                ref: attrs.ref,
                dt2D: attrs.dt2D,
                dtr: attrs.dtr,
                r1: attrs.r1,
                r2: attrs.r2,
              };
            }
            break;
          case 'v':
          case 't':
            capture = true;
            break;
          case 'oddHeader':
          case 'oddFooter':
          case 'evenHeader':
          case 'evenFooter':
          case 'firstHeader':
          case 'firstFooter':
            // A `<headerFooter>` child carries its header/footer definition as text (the `&`-prefixed
            // section/format tokens, e.g. `&C&"Arial"&G`). Capture it verbatim so a round-trip preserves
            // a header image's `&G` picture token and every other formatting directive.
            capture = true;
            break;
          case 'mergeCell':
            // A well-formed file never declares overlapping merges; a corrupt one might. Reject the
            // bad range at the model boundary, but don't let one abort the whole parse — drop it and
            // keep reading the valid geometry.
            if (attrs.ref !== undefined && attrs.ref !== '') {
              try {
                sheet.mergeCells(attrs.ref);
              } catch {
                // overlapping/malformed merge in the source file — skip it
              }
            }
            break;
          case 'tabColor':
            // A self-closing `<sheetPr>` child, so it arrives here rather than as text.
            sheet.tabColor = parseColor(attrs);
            break;
          case 'outlinePr':
            // Another self-closing `<sheetPr>` child; only set flags the source actually carried, so
            // a file without them leaves `outline` empty and a re-write stays byte-clean.
            if (attrs.summaryBelow !== undefined)
              sheet.outline.summaryBelow = boolPresent(attrs.summaryBelow);
            if (attrs.summaryRight !== undefined)
              sheet.outline.summaryRight = boolPresent(attrs.summaryRight);
            break;
          case 'pane':
            // A `<sheetView>` child recording a frozen (or split) pane. Only a frozen pane maps onto
            // the model's view; a source without one leaves `view` empty, so a re-write emits no pane.
            if (attrs.state === 'frozen' || attrs.state === 'frozenSplit') {
              sheet.view.state = 'frozen';
              if (attrs.xSplit !== undefined) sheet.view.xSplit = Number(attrs.xSplit);
              if (attrs.ySplit !== undefined) sheet.view.ySplit = Number(attrs.ySplit);
              if (attrs.topLeftCell !== undefined) sheet.view.topLeftCell = attrs.topLeftCell;
            }
            break;
          case 'pageSetUpPr':
            // The fit-to-page flag, a self-closing `<sheetPr>` child. Recorded only when the source
            // carried the attribute, so a `<pageSetUpPr>` present for other reasons (e.g.
            // `autoPageBreaks`) leaves `pageSetup.fitToPage` unset.
            if (attrs.fitToPage !== undefined)
              sheet.pageSetup.fitToPage = boolPresent(attrs.fitToPage);
            break;
          case 'printOptions':
            applyPrintOptions(sheet.printOptions, attrs);
            break;
          case 'pageMargins':
            applyMargins(sheet.pageMargins, attrs);
            break;
          case 'pageSetup':
            applyPageSetup(sheet.pageSetup, attrs);
            break;
          case 'rowBreaks':
            pageBreaks.begin(sheet.rowBreaks);
            break;
          case 'colBreaks':
            pageBreaks.begin(sheet.columnBreaks);
            break;
          case 'brk':
            pageBreaks.add(attrs);
            break;
          case 'sheetProtection': {
            const protection = parseSheetProtection(attrs);
            if (protection !== undefined) sheet.restoreProtection(protection);
            break;
          }
          case 'autoFilter':
            autoFilter.begin(attrs);
            break;
          case 'filterColumn':
            autoFilter.beginColumn(attrs);
            break;
          case 'filters':
            autoFilter.beginValues(attrs);
            break;
          case 'filter':
            autoFilter.addValue(attrs);
            break;
          case 'customFilters':
            autoFilter.beginCustom(attrs);
            break;
          case 'customFilter':
            autoFilter.addCustom(attrs);
            break;
          default:
            // A run's `<rPr>` child (`<b/>`, `<sz>`, `<color>`, `<rFont>`, …) sets one font facet; it
            // is self-closing, so it is read here on open. Nothing else uses the default branch.
            runs.applyProperty(local, attrs);
            break;
        }
        if (selfClosing && (local === 'f' || local === 'v')) capture = false;
      },
      onText(chunk) {
        if (capture) text += chunk;
      },
      onClose(name) {
        const local = localName(name);
        switch (local) {
          case 'f':
            formula = text;
            hasFormula = true;
            break;
          case 'v':
            valueText = text;
            hasValue = true;
            break;
          case 't':
            // A `<t>` inside a run is that run's text; a bare `<t>` directly in the `<is>` is a plain
            // inline string. A run takes precedence — a run is also inside the inline string.
            if (!runs.appendText(text) && inInlineString) inlineText += text;
            break;
          case 'r':
            runs.endRun();
            break;
          case 'is':
            inInlineString = false;
            break;
          case 'oddHeader':
          case 'oddFooter':
          case 'evenHeader':
          case 'evenFooter':
          case 'firstHeader':
          case 'firstFooter':
            sheet.headerFooter[local] = text;
            break;
          case 'c':
            finalizeCellFromState();
            break;
          case 'row':
            rowStyle = -1;
            rowCustomFormat = false;
            break;
          case 'filterColumn':
            autoFilter.endColumn();
            break;
          case 'autoFilter':
            autoFilter.commit(sheet);
            break;
          case 'rowBreaks':
          case 'colBreaks':
            pageBreaks.end();
            break;
          default:
            break;
        }
        capture = false;
      },
    },
    {closeEmptyElements: WORKSHEET_EMPTY_CLOSES},
  );
}

function applyColumn(
  sheet: Worksheet,
  attrs: {readonly [k: string]: string},
  xfStyles: ReadonlyArray<XfStyle>,
  columnStyle: Map<number, number>,
): void {
  const min = Number(attrs.min);
  const max = Number(attrs.max);
  if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1) return;
  const width = attrs.width !== undefined ? Number(attrs.width) : undefined;
  const hidden = boolStrict(attrs.hidden);
  const styleIndex = attrs.style !== undefined ? Number(attrs.style) : -1;
  // The column's style resolves to the same facet bundle a cell's does; mirror all of it onto the
  // column model so `getColumn(i)` reflects the declared default, not just its number format.
  const style = styleIndex >= 0 ? xfStyles[styleIndex] : undefined;
  for (let index = min; index <= max; index++) {
    const properties = sheet.getColumn(index);
    if (width !== undefined && Number.isFinite(width) && attrs.customWidth !== '0')
      properties.width = width;
    if (hidden) properties.hidden = true;
    if (attrs.outlineLevel !== undefined) {
      const level = Number(attrs.outlineLevel);
      if (Number.isInteger(level) && level > 0) properties.outlineLevel = level;
    }
    if (boolStrict(attrs.collapsed)) properties.collapsed = true;
    if (style?.numFmt !== undefined) properties.numFmt = style.numFmt;
    if (style?.fill !== undefined) properties.fill = style.fill;
    if (style?.font !== undefined) properties.font = style.font;
    if (style?.border !== undefined) properties.border = style.border;
    if (style?.alignment !== undefined) properties.alignment = style.alignment;
    if (style?.protection !== undefined) properties.protection = style.protection;
    // Record the column's style so a bare cell in it can inherit the full column format on read.
    if (styleIndex >= 0) columnStyle.set(index, styleIndex);
  }
}

function applyRow(sheet: Worksheet, attrs: {readonly [k: string]: string}): void {
  const number = Number(attrs.r);
  if (!Number.isInteger(number) || number < 1) return;
  const properties = sheet.getRow(number);
  if (attrs.ht !== undefined && attrs.customHeight !== '0') {
    const height = Number(attrs.ht);
    if (Number.isFinite(height)) properties.height = height;
  }
  if (boolStrict(attrs.hidden)) properties.hidden = true;
  if (attrs.outlineLevel !== undefined) {
    const level = Number(attrs.outlineLevel);
    if (Number.isInteger(level) && level > 0) properties.outlineLevel = level;
  }
  if (boolStrict(attrs.collapsed)) properties.collapsed = true;
}

// Read the `<printOptions>` boolean toggles back onto the model, storing only the ones the source
// carried so a re-write stays byte-clean. An OOXML boolean is `1`/`true` for on and `0`/`false` for
// off; a present-but-unrecognised token is dropped rather than coerced.
function applyPrintOptions(
  printOptions: PrintOptions,
  attrs: {readonly [k: string]: string},
): void {
  const horizontalCentered = boolTristate(attrs.horizontalCentered);
  if (horizontalCentered !== undefined) printOptions.horizontalCentered = horizontalCentered;
  const verticalCentered = boolTristate(attrs.verticalCentered);
  if (verticalCentered !== undefined) printOptions.verticalCentered = verticalCentered;
  const headings = boolTristate(attrs.headings);
  if (headings !== undefined) printOptions.headings = headings;
  const gridLines = boolTristate(attrs.gridLines);
  if (gridLines !== undefined) printOptions.gridLines = gridLines;
  const gridLinesSet = boolTristate(attrs.gridLinesSet);
  if (gridLinesSet !== undefined) printOptions.gridLinesSet = gridLinesSet;
}

function applyMargins(margins: PageMargins, attrs: {readonly [k: string]: string}): void {
  for (const side of MARGIN_SIDES) {
    const raw = attrs[side];
    if (raw === undefined) continue;
    const value = Number(raw);
    if (Number.isFinite(value)) margins[side] = value;
  }
}

// Read the `<pageSetup>` print-scaling attributes back onto the model, setting only those the
// source carried so a re-write stays byte-clean. Numeric attributes that fail to parse are
// dropped rather than stored as NaN; the enumerated ones are trusted verbatim (an unexpected token
// round-trips harmlessly as an unknown string).
function applyPageSetup(pageSetup: PageSetup, attrs: {readonly [k: string]: string}): void {
  const num = (raw: string | undefined): number | undefined => {
    if (raw === undefined) return undefined;
    const value = Number(raw);
    return Number.isFinite(value) ? value : undefined;
  };
  const paperSize = num(attrs.paperSize);
  if (paperSize !== undefined) pageSetup.paperSize = paperSize;
  const scale = num(attrs.scale);
  if (scale !== undefined) pageSetup.scale = scale;
  const fitToWidth = num(attrs.fitToWidth);
  if (fitToWidth !== undefined) pageSetup.fitToWidth = fitToWidth;
  const fitToHeight = num(attrs.fitToHeight);
  if (fitToHeight !== undefined) pageSetup.fitToHeight = fitToHeight;
  if (attrs.pageOrder === 'downThenOver' || attrs.pageOrder === 'overThenDown') {
    pageSetup.pageOrder = attrs.pageOrder;
  }
  if (attrs.orientation === 'portrait' || attrs.orientation === 'landscape') {
    pageSetup.orientation = attrs.orientation;
  }
}

function finalizeCell(
  sheet: Worksheet,
  ref: string,
  raw: RawCell,
  sharedStrings: readonly SharedString[],
  style: XfStyle | undefined,
): void {
  const {col, row} = decodeAddress(ref);
  if (col === undefined || row === undefined) return;
  const cell = sheet.getCell(ref);
  applyCellStyle(cell, style);
  cell.value = decodeCellContent(raw, sharedStrings, style?.numFmt);
}

// Applies a resolved xf's non-value facets to a cell. Shared by the ordinary cell path and the
// shared-formula clone path, so a styled clone (fill/font/border/alignment/protection) keeps its
// look on read rather than surviving as value-only.
function applyCellStyle(cell: Cell, style: XfStyle | undefined): void {
  if (style?.fill !== undefined) cell.fill = style.fill;
  if (style?.numFmt !== undefined) cell.numFmt = style.numFmt;
  if (style?.font !== undefined) cell.font = style.font;
  if (style?.border !== undefined) cell.border = style.border;
  if (style?.alignment !== undefined) cell.alignment = style.alignment;
  if (style?.protection !== undefined) cell.protection = style.protection;
  if (style?.quotePrefix !== undefined) cell.quotePrefix = style.quotePrefix;
  if (style?.xfId !== undefined) cell.namedStyleId = style.xfId;
}
