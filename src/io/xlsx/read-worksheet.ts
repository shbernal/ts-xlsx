// The worksheet-body reader: one `xl/worksheets/sheetN.xml` in, its cells/columns/rows/print-setup
// applied onto a {@link Worksheet}. It is a single streaming pass driving a handful of accumulators
// (the cell being read, shared-formula masters, an autofilter draft, the current page-break axis) so
// each element commits its state as it closes. Style indices resolve through the parsed style table.

import {MAX_COLUMN, MAX_ROW, tryDecodeRange} from '../../core/address.ts';
import {
  type CustomFilterPredicate,
  type FilterColumn,
  type FilterCriteria,
  isCustomFilterOperator,
} from '../../core/autofilter.ts';
import {INTERNAL} from '../../core/internal.ts';
import {
  isPageOrder,
  isPageOrientation,
  type PageBreak,
  type PageMargins,
  type PageSetup,
  type PrintOptions,
} from '../../core/page-setup.ts';
import {
  SHEET_PROTECTION_FLAGS,
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionFlags,
} from '../../core/protection.ts';
import {assignStyleFacets} from '../../core/style.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {
  boolPresent,
  boolStrict,
  boolTristate,
  decodeSpreadsheetText,
  enumToken,
  localName,
  numFinite,
  numInteger,
  parseXml,
  type XmlAttributes,
} from '../../xml/xml-read.ts';
import type {XfStyle} from '../style/xf-style.ts';
import {CellAccumulator} from './cell-accumulator.ts';
import type {SharedString} from './cell-value.ts';
import {parseColor} from './color-xml.ts';

const MARGIN_SIDES = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const;

// Worksheet elements that commit on their close: a formatted-but-empty `<c/>` and a criteria-free
// self-closing `<autoFilter/>` are expanded to open+close so each finalises once in onClose. The
// text-bearing `<f/>`/`<v/>`/`<t/>` are deliberately excluded: an empty one must not commit.
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

// Read a <sheetProtection> element back into a SheetProtection: the deserialization mirror of the
// writer. `sheet="0"` (or "false") means the element records an *un*protected sheet, so nothing is
// restored. Each flag attribute is the INVERSE of the author's allow-flag ("1" forbids, "0" permits),
// and only attributes actually present are carried, so an omitted (default-valued) flag stays absent,
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
  const spin = numInteger(spinCount, 0);
  if (
    algorithmName !== undefined &&
    hashValue !== undefined &&
    saltValue !== undefined &&
    spin !== undefined
  ) {
    const credential: SheetProtectionCredential = {
      algorithmName,
      hashValue,
      saltValue,
      spinCount: spin,
    };
    return {flags, credential};
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
    this.#colId = numInteger(attrs.colId, 0) ?? -1;
    this.#values = null;
    this.#blank = false;
    this.#predicates = null;
    this.#and = false;
  }

  beginValues(attrs: XmlAttributes): void {
    this.#values = [];
    // `blank` defaults off when absent, so presence is required first; then it reads as an
    // on-when-present flag. (`boolPresent` alone would treat the absent attribute as on.)
    this.#blank = attrs.blank !== undefined && boolPresent(attrs.blank);
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
  // filterable and is dropped so a re-write stays clean: load-repair, not authoring.
  endColumn(): void {
    const criteria = pendingFilterCriteria(this.#values, this.#blank, this.#predicates, this.#and);
    if (this.#colId >= 0 && criteria !== null) {
      this.#columns.push({colId: this.#colId, criteria});
    }
  }

  // Commit the accumulated autofilter to the sheet. Runs on `</autoFilter>`, including the synthesized
  // close of a criteria-free self-closing `<autoFilter/>`. Columns whose colId falls outside the range
  // are dropped here so the strict setter never trips on hostile input.
  commit(sheet: Worksheet): void {
    if (this.#ref === null) return;
    const decoded = tryDecodeRange(this.#ref);
    const left = decoded?.left;
    const right = decoded?.right;
    // A filter needs a bounded rectangle, so an unbounded or unreadable range leaves the sheet
    // without one; `canonicalizeAutoFilter` would refuse it anyway, and refusing here keeps the
    // authoring guard a guard rather than a control-flow path.
    if (left !== undefined && right !== undefined) {
      const width = right - left + 1;
      sheet.autoFilter = {ref: this.#ref, columns: this.#columns.filter((c) => c.colId < width)};
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
    const id = numInteger(attrs.id, 1);
    if (id === undefined) return;
    const brk: {id: number; max?: number; man?: boolean} = {id};
    const max = numInteger(attrs.max, 0);
    if (max !== undefined) brk.max = max;
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
  const cell = new CellAccumulator({richRuns: true});
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
  // (when customFormat), then its column's default: the order Excel applies. Runs on `</c>` close,
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
        if (cell.openElement(local, attrs, selfClosing)) return;
        switch (local) {
          case 'col':
            applyColumn(sheet, attrs, xfStyles, columnStyle);
            break;
          case 'row':
            applyRow(sheet, attrs);
            rowStyle = numInteger(attrs.s, 0) ?? -1;
            rowCustomFormat = boolStrict(attrs.customFormat);
            break;
          case 'oddHeader':
          case 'oddFooter':
          case 'evenHeader':
          case 'evenFooter':
          case 'firstHeader':
          case 'firstFooter':
            // A `<headerFooter>` child carries its header/footer definition as text (the `&`-prefixed
            // section/format tokens, e.g. `&C&"Arial"&G`). Capture the whole of it so a round-trip
            // preserves a header image's `&G` picture token and every other formatting directive.
            cell.capture();
            break;
          case 'mergeCell':
            // A well-formed file never declares overlapping merges; a corrupt one might. Reject the
            // bad range at the model boundary, but don't let one abort the whole parse: drop it and
            // keep reading the valid geometry.
            if (attrs.ref !== undefined && attrs.ref !== '') {
              try {
                sheet.mergeCells(attrs.ref);
              } catch {
                // overlapping/malformed merge in the source file: skip it
              }
            }
            break;
          case 'tabColor':
          case 'outlinePr':
          case 'sheetView':
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
            if (protection !== undefined) sheet[INTERNAL].restoreProtection(protection);
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
        }
      },
      onText(chunk) {
        cell.appendChunk(chunk);
      },
      onClose(name) {
        const local = localName(name);
        const claimed = cell.closeElement(local);
        if (claimed === 'cell') {
          finalizeCellFromState();
          return;
        }
        if (claimed === 'claimed') return;
        switch (local) {
          case 'oddHeader':
          case 'oddFooter':
          case 'evenHeader':
          case 'evenFooter':
          case 'firstHeader':
          case 'firstFooter':
            // Header text carries the `_xHHHH_` convention, same as a cell value: Excel decodes it
            // here and re-emits it on save (measured: a patched `_x0001_` reads back over COM as
            // U+0001, and a `_x005F_x0041_` as the literal `_x0041_`). The decode is on the whole
            // element text, never on a SAX chunk. See {@link decodeSpreadsheetText}.
            sheet.headerFooter[local] = decodeSpreadsheetText(cell.capturedText);
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
      },
    },
    {closeEmptyElements: WORKSHEET_EMPTY_CLOSES},
  );
}

// Apply one `<sheetPr>` / `<sheetView>` / print-setup child to the sheet. These are the worksheet's
// layout and print metadata; grouping them here keeps the cell-reading switch a pure dispatch. Each
// records only what the source carried, so a file missing a facet leaves it unset and a re-write
// stays byte-clean. Each is read on open, from its attributes alone, which is why `<sheetView>`
// belongs here despite wrapping children: what this reads of it is attributes, and its `<pane>`
// child arrives as its own dispatch.
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
    case 'sheetView':
      // Excel omits `showGridLines` when the grid is on, so only a present-and-false attribute is
      // recorded. Leaving it unset otherwise is what keeps a re-write from fabricating the
      // attribute on every sheet that never mentioned it.
      if (attrs.showGridLines !== undefined && !boolPresent(attrs.showGridLines)) {
        sheet.view.showGridLines = false;
      }
      break;
    case 'pane':
      // A `<sheetView>` child recording a frozen (or split) pane. Only a frozen pane maps onto the
      // model's view; a source without one leaves `view` empty, so a re-write emits no pane.
      if (attrs.state === 'frozen' || attrs.state === 'frozenSplit') {
        sheet.view.state = 'frozen';
        // A split that is not a non-negative integer is dropped, not stored: `Worksheet.freeze()`
        // refuses the same value, and storing it here only defers the failure to the writer, which
        // adds the split to a cell ordinal and throws about a column the caller never named.
        const xSplit = numInteger(attrs.xSplit, 0);
        const ySplit = numInteger(attrs.ySplit, 0);
        if (xSplit !== undefined) sheet.view.xSplit = xSplit;
        if (ySplit !== undefined) sheet.view.ySplit = ySplit;
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
  const min = numInteger(attrs.min, 1);
  const max = numInteger(attrs.max, 1);
  if (min === undefined || max === undefined || min > MAX_COLUMN) return;
  // Clamp the span to the format's ceiling rather than letting `getColumn` throw through the read:
  // a `<col max="99999999">` is a file Excel opens, and an unclamped loop would materialise 16.7
  // million column records before dying. Same reading as the streaming reader's `collectHiddenColumn`.
  const last = Math.min(max, MAX_COLUMN);
  const width = numFinite(attrs.width);
  const hidden = boolStrict(attrs.hidden);
  const styleIndex = numInteger(attrs.style, 0) ?? -1;
  // The column's style resolves to the same facet bundle a cell's does; mirror all of it onto the
  // column model so `getColumn(i)` reflects the declared default, not just its number format.
  const style = styleIndex >= 0 ? xfStyles[styleIndex] : undefined;
  for (let index = min; index <= last; index++) {
    const column = sheet.getColumn(index);
    if (width !== undefined && boolPresent(attrs.customWidth)) column.width = width;
    if (hidden) column.hidden = true;
    const outlineLevel = numInteger(attrs.outlineLevel, 1);
    if (outlineLevel !== undefined) column.outlineLevel = outlineLevel;
    if (boolStrict(attrs.collapsed)) column.collapsed = true;
    if (style !== undefined) assignStyleFacets(column, style);
    // Record the column's style so a bare cell in it can inherit the full column format on read.
    if (styleIndex >= 0) columnStyle.set(index, styleIndex);
  }
}

function applyRow(sheet: Worksheet, attrs: XmlAttributes): void {
  const number = numInteger(attrs.r, 1);
  // Out-of-grid rows are dropped rather than clamped: unlike a `<col>` span, an `<r>` names one row,
  // so there is nothing to fold it onto and clamping would silently move its formatting to 1048576.
  if (number === undefined || number > MAX_ROW) return;
  // A `<row>` that states no attribute at all leaves no format record behind: the handle creates
  // one only when something is written through it. That is the right reading: a bare `<row r="5"/>`
  // carries no formatting to round-trip, and fabricating an empty record for it would put row 5 in
  // the used range on the strength of an element that says nothing.
  const row = sheet.getRow(number);
  const height = numFinite(attrs.ht);
  if (height !== undefined && boolPresent(attrs.customHeight)) row.height = height;
  if (boolStrict(attrs.hidden)) row.hidden = true;
  const outlineLevel = numInteger(attrs.outlineLevel, 1);
  if (outlineLevel !== undefined) row.outlineLevel = outlineLevel;
  if (boolStrict(attrs.collapsed)) row.collapsed = true;
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
    const value = numFinite(attrs[side]);
    if (value !== undefined) margins[side] = value;
  }
}

// Read the `<pageSetup>` print-scaling attributes back onto the model, setting only those the
// source carried so a re-write stays byte-clean. Each of the four is a count or an enumeration id,
// so a fractional or negative one carries no meaning and is dropped; the enumerated string
// attributes are trusted verbatim (an unexpected token round-trips harmlessly as an unknown string).
function applyPageSetup(pageSetup: PageSetup, attrs: XmlAttributes): void {
  const paperSize = numInteger(attrs.paperSize, 0);
  if (paperSize !== undefined) pageSetup.paperSize = paperSize;
  const scale = numInteger(attrs.scale, 0);
  if (scale !== undefined) pageSetup.scale = scale;
  const fitToWidth = numInteger(attrs.fitToWidth, 0);
  if (fitToWidth !== undefined) pageSetup.fitToWidth = fitToWidth;
  const fitToHeight = numInteger(attrs.fitToHeight, 0);
  if (fitToHeight !== undefined) pageSetup.fitToHeight = fitToHeight;
  const pageOrder = enumToken(attrs.pageOrder, isPageOrder);
  if (pageOrder !== undefined) pageSetup.pageOrder = pageOrder;
  const orientation = enumToken(attrs.orientation, isPageOrientation);
  if (orientation !== undefined) pageSetup.orientation = orientation;
}
