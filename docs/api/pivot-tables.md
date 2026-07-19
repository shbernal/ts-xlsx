# Pivot tables

<!-- Generated from the public types by `npm run docs`. Do not edit by hand. -->

### `PivotMetric`

<sub>type</sub>

The aggregation a pivot's value field applies. These are OOXML's `ST_DataConsolidateFunction`
names verbatim, so a metric doubles as its `<dataField subtotal="…">` value. Excel performs the
aggregation itself on refresh; the writer only records which function to apply.

```ts
type PivotMetric = 'sum' | 'count' | 'countNums' | 'average' | 'max' | 'min' | 'product' | 'stdDev' | 'stdDevp' | 'var' | 'varp';
```

---

### `PivotTable`

<sub>class</sub>

A pivot table built over a source sheet's data. Construction reads the source once and computes
the full cache (fields + records) and the axis-field wiring the renderer needs; nothing here
touches XML.

Supported shape: exactly one value field aggregated by `sum`, at least one row field and one
column field. An unsupported request throws at authoring time rather than emitting a corrupt file.

```ts
class PivotTable {
  readonly metric: PivotMetric;
  readonly sourceSheetName: string;
  readonly sourceRef: string;
  readonly cacheFields: readonly PivotCacheField[];
  readonly records: readonly (readonly PivotRecordCell[])[];
  readonly rowFields: readonly number[];
  readonly columnFields: readonly number[];
  readonly valueField: number;
  get valueFieldName(): string;
}
```

**Members**

- `readonly sourceRef: string;` — The `A1:C4` source range: the header row through the last data row, across the field columns.
- `readonly rowFields: readonly number[];` — Indices into `cacheFields` of the row-axis, column-axis, and value fields.
- `get valueFieldName(): string;` — The value field's header name, used to label the aggregated data column ("Sum of Amount").

---

### `PivotTableOptions`

<sub>interface</sub>

How a pivot table is authored: a source sheet and the header names that drive each axis.
`rows`/`columns`/`values` name columns by their header text in the source's first row.

```ts
interface PivotTableOptions {
    readonly source: Worksheet;
    readonly rows: readonly string[];
    readonly columns: readonly string[];
    readonly values: readonly string[];
    readonly metric?: PivotMetric;
}
```
