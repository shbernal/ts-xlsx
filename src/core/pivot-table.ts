// A pivot table authored from a model.
//
// A pivot summarises a source range: its distinct field values become row/column axes and a value
// field is aggregated across them. OOXML splits that into three parts — a `pivotCacheDefinition`
// (the field catalogue), a `pivotCacheRecords` (a copy of the source rows, with axis-field cells
// swapped for indices into the catalogue), and a `pivotTableDefinition` (the layout on the
// destination sheet). This module owns the *semantic* computation of all three; the OOXML rendering
// lives in `io/xlsx/pivot.ts`.
//
// The source data is captured when the pivot is added — the model reads the source sheet's cells
// once, here, so the pivot is a stable snapshot independent of later edits to the source.

import {encodeAddress} from './address.ts';
import {
  type CellValue,
  isErrorValue,
  isFormulaValue,
  isHyperlinkValue,
  isRichTextValue,
  isSharedFormulaValue,
  type RichTextValue,
} from './value.ts';
import type {Worksheet} from './worksheet.ts';

/** The aggregation a pivot's value field applies. These are OOXML's `ST_DataConsolidateFunction`
 * names verbatim, so a metric doubles as its `<dataField subtotal="…">` value. Excel performs the
 * aggregation itself on refresh; the writer only records which function to apply. */
export type PivotMetric =
  | 'sum'
  | 'count'
  | 'countNums'
  | 'average'
  | 'max'
  | 'min'
  | 'product'
  | 'stdDev'
  | 'stdDevp'
  | 'var'
  | 'varp';

const PIVOT_METRICS: ReadonlySet<PivotMetric> = new Set<PivotMetric>([
  'sum',
  'count',
  'countNums',
  'average',
  'max',
  'min',
  'product',
  'stdDev',
  'stdDevp',
  'var',
  'varp',
]);

/** Map an OOXML `<dataField subtotal="…">` value back to its metric. The attribute is absent for
 * `sum` (Excel's implicit default), so `undefined` reads as `sum`; an unrecognised value also reads
 * as `sum` rather than throwing, because reconstructing an existing file is a lenient operation —
 * the strict rejection of unknown metrics belongs on the authoring path, not the read path. */
export function pivotMetricFromSubtotal(subtotal: string | undefined): PivotMetric {
  if (subtotal === undefined) return 'sum';
  return PIVOT_METRICS.has(subtotal as PivotMetric) ? (subtotal as PivotMetric) : 'sum';
}

/** One field in a loaded pivot's cache catalogue, in declared order; the pivot refers to it by index. */
export interface ParsedPivotField {
  readonly name: string;
}

/** Where a worksheet-backed pivot cache draws its rows from. Empty strings stand in for a cache whose
 * source is not a worksheet (an external or consolidation source), which the reader does not model. */
export interface ParsedPivotSource {
  readonly sheet: string;
  readonly ref: string;
}

/** The semantic model reconstructed from a loaded pivot's `pivotTableDefinition` and its
 * `pivotCacheDefinition` (see `io/xlsx/pivot-read.ts`). Field roles are indices into {@link fields};
 * {@link metric} is the aggregation the value field applies. This mirrors the authoring model's shape
 * without requiring the source sheet it was built from, so a pivot loaded from a package is
 * inspectable data rather than an opaque preserved blob. It is a read-only view: the writer emits a
 * loaded pivot from its preserved parts, not from this model, so exposing it never double-emits. */
export interface ParsedPivotTable {
  readonly name: string;
  readonly cacheId: string;
  readonly source: ParsedPivotSource;
  readonly fields: readonly ParsedPivotField[];
  readonly rowFields: readonly number[];
  readonly columnFields: readonly number[];
  /** Index into {@link fields} of the aggregated field, or -1 when no `<dataField>` was declared. */
  readonly valueField: number;
  readonly valueFieldName: string;
  /** The `<dataField>`'s own caption ("Average of Amount"), which Excel shows on the data column. */
  readonly valueCaption: string;
  readonly metric: PivotMetric;
}

/** How a pivot table is authored: a source sheet and the header names that drive each axis.
 * `rows`/`columns`/`values` name columns by their header text in the source's first row. */
export interface PivotTableOptions {
  readonly source: Worksheet;
  readonly rows: readonly string[];
  readonly columns: readonly string[];
  readonly values: readonly string[];
  readonly metric?: PivotMetric;
}

/** One distinct value in a cache field's shared-items catalogue, or an inline record cell. A
 * `blank` is a missing source value, serialised as `<m/>` rather than an empty string. */
export type PivotItem =
  | {readonly kind: 'string'; readonly value: string}
  | {readonly kind: 'number'; readonly value: number}
  | {readonly kind: 'blank'};

/** The numeric summary Excel expects on a non-shared field whose every present value is a number. */
export interface PivotNumericSummary {
  readonly allInteger: boolean;
  readonly min: number;
  readonly max: number;
}

/** One field of the pivot cache. An axis field (row or column) carries a `sharedItems` catalogue its
 * records reference by index; any other field stores its values inline in the records and, when they
 * are all numeric, describes them with a `numeric` summary. */
export interface PivotCacheField {
  readonly name: string;
  readonly sharedItems: readonly PivotItem[] | null;
  readonly numeric: PivotNumericSummary | null;
  readonly containsBlank: boolean;
}

/** One cell of a cache record: an index into a shared-items catalogue, or an inline value. */
export type PivotRecordCell = {readonly kind: 'index'; readonly index: number} | PivotItem;

const BLANK: PivotItem = {kind: 'blank'};

/**
 * A pivot table built over a source sheet's data. Construction reads the source once and computes
 * the full cache (fields + records) and the axis-field wiring the renderer needs; nothing here
 * touches XML.
 *
 * Supported shape: exactly one value field aggregated by `sum`, at least one row field and one
 * column field. An unsupported request throws at authoring time rather than emitting a corrupt file.
 */
export class PivotTable {
  readonly metric: PivotMetric;
  readonly sourceSheetName: string;
  /** The `A1:C4` source range: the header row through the last data row, across the field columns. */
  readonly sourceRef: string;
  readonly cacheFields: readonly PivotCacheField[];
  readonly records: readonly (readonly PivotRecordCell[])[];
  /** Indices into {@link cacheFields} of the row-axis, column-axis, and value fields. */
  readonly rowFields: readonly number[];
  readonly columnFields: readonly number[];
  readonly valueField: number;

  constructor(options: PivotTableOptions) {
    const metric = options.metric ?? 'sum';
    if (!PIVOT_METRICS.has(metric)) {
      throw new Error(`unsupported pivot metric "${metric}" — expected one of ${[...PIVOT_METRICS].join(', ')}`);
    }
    this.metric = metric;

    const source = options.source;
    const columnCount = source.columnCount;
    const lastRow = source.rowCount;
    if (columnCount < 1 || lastRow < 2) {
      throw new Error('a pivot source needs a header row and at least one data row');
    }

    // Every non-blank header cell in row 1 defines a field, in ascending column order.
    const fields: {readonly name: string; readonly col: number}[] = [];
    for (let col = 1; col <= columnCount; col++) {
      const name = textOf(scalarOf(source.getCell(encodeAddress(col, 1)).value));
      if (name !== '') fields.push({name, col});
    }
    if (fields.length === 0) throw new Error('the pivot source header row is empty');

    const resolve = (role: string, name: string): number => {
      const index = fields.findIndex(field => field.name === name);
      if (index < 0) {
        throw new Error(`pivot ${role} field "${name}" is not a column header in the source sheet`);
      }
      return index;
    };
    if (options.rows.length === 0) throw new Error('a pivot table needs at least one row field');
    if (options.columns.length === 0) throw new Error('a pivot table needs at least one column field');
    if (options.values.length !== 1) throw new Error('a pivot table needs exactly one value field');

    this.rowFields = options.rows.map(name => resolve('row', name));
    this.columnFields = options.columns.map(name => resolve('column', name));
    this.valueField = resolve('value', options.values[0] as string);
    const axisFields = new Set<number>([...this.rowFields, ...this.columnFields]);

    const firstCol = fields[0]?.col as number;
    const lastCol = fields[fields.length - 1]?.col as number;
    this.sourceSheetName = source.name;
    this.sourceRef = `${encodeAddress(firstCol, 1)}:${encodeAddress(lastCol, lastRow)}`;

    // Read the source body once, field by field, so the same scan feeds both the shared-items
    // catalogues and the records that reference them.
    const dataRowCount = lastRow - 1;
    const columnScalars = fields.map(field => {
      const scalars: PivotItem[] = [];
      for (let row = 2; row <= lastRow; row++) {
        scalars.push(scalarOf(source.getCell(encodeAddress(field.col, row)).value));
      }
      return scalars;
    });

    const catalogues: (Map<string, number> | null)[] = fields.map(() => null);
    this.cacheFields = fields.map((field, fieldIndex) => {
      const scalars = columnScalars[fieldIndex] as PivotItem[];
      const containsBlank = scalars.some(scalar => scalar.kind === 'blank');
      if (axisFields.has(fieldIndex)) {
        const items: PivotItem[] = [];
        const catalogue = new Map<string, number>();
        for (const scalar of scalars) {
          const key = itemKey(scalar);
          if (!catalogue.has(key)) {
            catalogue.set(key, items.length);
            items.push(scalar);
          }
        }
        catalogues[fieldIndex] = catalogue;
        return {name: field.name, sharedItems: items, numeric: null, containsBlank};
      }
      return {name: field.name, sharedItems: null, numeric: numericSummary(scalars), containsBlank};
    });

    const records: (readonly PivotRecordCell[])[] = [];
    for (let row = 0; row < dataRowCount; row++) {
      records.push(
        fields.map((_field, fieldIndex): PivotRecordCell => {
          const scalar = (columnScalars[fieldIndex] as PivotItem[])[row] as PivotItem;
          const catalogue = catalogues[fieldIndex];
          if (!catalogue) return scalar;
          return {kind: 'index', index: catalogue.get(itemKey(scalar)) as number};
        })
      );
    }
    this.records = records;
  }

  /** The value field's header name, used to label the aggregated data column ("Sum of Amount"). */
  get valueFieldName(): string {
    return (this.cacheFields[this.valueField] as PivotCacheField).name;
  }
}

/** A stable dedup key for a shared item: kind-tagged so the number `1` and the string `"1"` differ. */
function itemKey(item: PivotItem): string {
  switch (item.kind) {
    case 'string':
      return `s:${item.value}`;
    case 'number':
      return `n:${item.value}`;
    case 'blank':
      return 'b';
  }
}

/** The numeric summary for an inline field, or null when any present value is non-numeric (a string
 * present means the field is not a pure numeric column; blanks alone do not disqualify it). */
function numericSummary(scalars: readonly PivotItem[]): PivotNumericSummary | null {
  let min = Infinity;
  let max = -Infinity;
  let allInteger = true;
  let sawNumber = false;
  for (const scalar of scalars) {
    if (scalar.kind === 'blank') continue;
    if (scalar.kind !== 'number') return null;
    sawNumber = true;
    if (scalar.value < min) min = scalar.value;
    if (scalar.value > max) max = scalar.value;
    if (!Number.isInteger(scalar.value)) allInteger = false;
  }
  return sawNumber ? {allInteger, min, max} : null;
}

/** The string form of a shared item, used for header names (a blank header contributes no field). */
function textOf(item: PivotItem): string {
  return item.kind === 'blank' ? '' : String(item.value);
}

/**
 * Reduce any cell value to the scalar a pivot cache can hold: a number, a string, or a blank. Only
 * finite numbers stay numeric (a NaN would corrupt the cache); every other kind is flattened to its
 * displayed text so hostile or exotic source content can never throw or leak an object into the XML.
 */
function scalarOf(value: CellValue): PivotItem {
  if (value === null) return BLANK;
  switch (typeof value) {
    case 'number':
      return Number.isFinite(value) ? {kind: 'number', value} : BLANK;
    case 'string':
      return {kind: 'string', value};
    case 'boolean':
      return {kind: 'string', value: value ? 'TRUE' : 'FALSE'};
    default:
      break;
  }
  if (value instanceof Date) return {kind: 'string', value: value.toISOString()};
  if (isRichTextValue(value)) return {kind: 'string', value: richTextToString(value)};
  if (isHyperlinkValue(value)) {
    return {kind: 'string', value: typeof value.text === 'string' ? value.text : richTextToString(value.text)};
  }
  if (isErrorValue(value)) return {kind: 'string', value: value.error};
  if (isFormulaValue(value) || isSharedFormulaValue(value)) {
    return value.result === undefined ? BLANK : scalarOf(value.result as CellValue);
  }
  return BLANK;
}

function richTextToString(value: RichTextValue): string {
  return value.richText.map(run => run.text).join('');
}
