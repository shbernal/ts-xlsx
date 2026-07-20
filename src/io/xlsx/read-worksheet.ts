// The worksheet-body reader: one `xl/worksheets/sheetN.xml` in, its cells/columns/rows/print-setup
// applied onto a {@link Worksheet}. It is a single streaming pass driving a handful of accumulators
// (the cell being read, shared-formula masters, an autofilter draft, the current page-break axis) so
// each element commits its state as it closes. Style indices resolve through the parsed style table.

import {decodeRange} from '../../core/address.ts';
import {
  type CustomFilterPredicate,
  type FilterColumn,
  type FilterCriteria,
  isCustomFilterOperator,
} from '../../core/autofilter.ts';
import type {PageBreak, PageMargins, PageSetup, PrintOptions} from '../../core/page-setup.ts';
import {
  SHEET_PROTECTION_FLAGS,
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionFlags,
} from '../../core/protection.ts';
import {assignStyleFacets} from '../../core/style.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {CellAccumulator} from './cell-accumulator.ts';
import type {SharedString} from './cell-value.ts';
import type {XfStyle} from './read-styles.ts';
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
function parseSheetProtection(attrs: XmlAttributes): SheetProtection | undefined {
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
  // The one `<c>` currently being read: its address/type/style, formula, value, inline text, rich
  // runs, and the sheet-spanning shared-formula master map. Each `<c>` resets it and commits it.
  const cell = new CellAccumulator();
  let inInlineString = false;
  let capture = false;
  let text = '';
  // A row with customFormat="1" supplies a default style for its cells that carry no `s`.
  let rowStyle = -1;
  let rowCustomFormat = false;
  const autoFilter = new AutoFilterAccumulator();
  const pageBreaks = new PageBreakAccumulator();
  // A column's `style` is the default for its cells that carry no style of their own; this
  // maps a column index to that style index so a bare cell can inherit it (as Excel does,
  // without stamping every cell). Columns are parsed before any cell references them.
  const columnStyle = new Map<number, number>();

  // Commit the cell held in the accumulator, resolving its style from its own `s`, then its row's
  // (when customFormat), then its column's default — the order Excel applies. Runs on `</c>` close,
  // including the synthesized close of a self-closing `<c/>` formatted-but-empty cell.
  const finalizeCellFromState = (): void => {
    const styleIndex =
      cell.styleIndex >= 0
        ? cell.styleIndex
        : rowCustomFormat && rowStyle >= 0
          ? rowStyle
          : (columnStyle.get(cell.col) ?? -1);
    const style = styleIndex >= 0 ? xfStyles[styleIndex] : xfStyles[0];
    cell.finalize(sheet, sharedStrings, style);
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
            cell.beginCell(attrs);
            break;
          case 'is':
            inInlineString = true;
            cell.beginInlineString();
            break;
          case 'r':
            // A run inside a rich inline string. Its `<rPr>` (if any) and `<t>` follow.
            if (inInlineString) cell.runs.beginRun();
            break;
          case 'rPr':
            // The run's formatting bundle; its self-closing children stream into the default branch.
            cell.runs.beginProperties();
            break;
          case 'f':
            capture = true;
            cell.beginFormula(attrs, selfClosing);
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
          case 'outlinePr':
          case 'pane':
          case 'pageSetUpPr':
          case 'printOptions':
          case 'pageMargins':
          case 'pageSetup':
            applySheetProperties(local, attrs, sheet);
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
            cell.runs.applyProperty(local, attrs);
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
            cell.setFormula(text);
            break;
          case 'v':
            cell.setValue(text);
            break;
          case 't':
            // A `<t>` inside a run is that run's text; a bare `<t>` directly in the `<is>` is a plain
            // inline string. A run takes precedence — a run is also inside the inline string.
            cell.appendText(text, inInlineString);
            break;
          case 'r':
            cell.runs.endRun();
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

// Apply one `<sheetPr>` / `<sheetView>` / print-setup child to the sheet. These are the worksheet's
// layout and print metadata; grouping them here keeps the cell-reading switch a pure dispatch. Each
// records only what the source carried, so a file missing a facet leaves it unset and a re-write
// stays byte-clean. Every element here is self-closing (its state is all in attributes), so it is
// read on open.
function applySheetProperties(local: string, attrs: XmlAttributes, sheet: Worksheet): void {
  switch (local) {
    case 'tabColor':
      // A `<sheetPr>` child.
      sheet.tabColor = parseColor(attrs);
      break;
    case 'outlinePr':
      // A `<sheetPr>` child.
      if (attrs.summaryBelow !== undefined)
        sheet.outline.summaryBelow = boolPresent(attrs.summaryBelow);
      if (attrs.summaryRight !== undefined)
        sheet.outline.summaryRight = boolPresent(attrs.summaryRight);
      break;
    case 'pane':
      // A `<sheetView>` child recording a frozen (or split) pane. Only a frozen pane maps onto the
      // model's view; a source without one leaves `view` empty, so a re-write emits no pane.
      if (attrs.state === 'frozen' || attrs.state === 'frozenSplit') {
        sheet.view.state = 'frozen';
        if (attrs.xSplit !== undefined) sheet.view.xSplit = Number(attrs.xSplit);
        if (attrs.ySplit !== undefined) sheet.view.ySplit = Number(attrs.ySplit);
        if (attrs.topLeftCell !== undefined) sheet.view.topLeftCell = attrs.topLeftCell;
      }
      break;
    case 'pageSetUpPr':
      // The fit-to-page flag, a `<sheetPr>` child. Recorded only when the attribute is present, so a
      // `<pageSetUpPr>` present for other reasons (e.g. `autoPageBreaks`) leaves `fitToPage` unset.
      if (attrs.fitToPage !== undefined) sheet.pageSetup.fitToPage = boolPresent(attrs.fitToPage);
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
  }
}

function applyColumn(
  sheet: Worksheet,
  attrs: XmlAttributes,
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
    if (style !== undefined) assignStyleFacets(properties, style);
    // Record the column's style so a bare cell in it can inherit the full column format on read.
    if (styleIndex >= 0) columnStyle.set(index, styleIndex);
  }
}

function applyRow(sheet: Worksheet, attrs: XmlAttributes): void {
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
function applyPrintOptions(printOptions: PrintOptions, attrs: XmlAttributes): void {
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

function applyMargins(margins: PageMargins, attrs: XmlAttributes): void {
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
function applyPageSetup(pageSetup: PageSetup, attrs: XmlAttributes): void {
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
