# getSheetValues() typescript definition is incorrect

## Desired behavior

The worksheet accessor that returns "all rows as a sparse array" must have a return type that precisely matches its runtime output. It does **not** return `Row` objects — it returns each row's *values*.

Each element is a row's value payload, which in this model can be either:

- a sparse, 1-indexed array of cell values (index `0` is empty, index `n` = value of column `n`), or
- a keyed object `{ [columnKey: string]: CellValue }` when the row was populated via column keys.

The outer array is itself sparse: it is 1-indexed by row number, with holes for empty/absent rows.

### Prior art / prior discussion

- Original declaration typed the result as `Row[]`, which is wrong — callers would reach for `Row` members that do not exist at runtime.
- One suggested fix was `Cell[][]`; also wrong, since elements are raw cell *values*, not `Cell` objects, and a row can be a keyed object rather than an array.
- Another suggested `any[]`, which is rejected on principle in this project — `any` erases the contract.
- The accurate shape mirrors the existing row-values union used elsewhere in the model: `type RowValues = CellValue[] | { [columnKey: string]: CellValue }`, giving a result of `RowValues[]` (sparse).

### Recommended shape for ts-xlsx

- Introduce/reuse a named `RowValues` (or equivalently named) union: `CellValue[] | { [columnKey: string]: CellValue }`.
- Type the accessor's return as a sparse array of `RowValues` (both the outer array and inner arrays are sparse/1-indexed; document the index-0 hole).
- Cover with a type-level test asserting the return type is assignable to/from the sparse `RowValues[]` shape and is NOT `Row[]`.

### The same union types `row.values` and `column.values`, not just the sheet accessor

The sparse `RowValues` shape is not confined to the whole-sheet accessor. The per-`Row` `values`
property is the same 1-indexed sparse array (or keyed object), and the per-`Column` `values`
property is the column's cells read top-to-bottom as a sparse array — both are `RowValues`-shaped
collections of `CellValue`, **never a scalar**. A published declaration once typed `Column.values`
as `string` (the result of a refactor that meant only to drop a `readonly` modifier), which is
doubly wrong: it is an array, not a string, and its elements are `CellValue`, not text. A caller
indexing `column.values[3]` then got a `string` character instead of the cell value, or a type error.

- `Row.values` and `Column.values` are typed as the shared sparse `RowValues`
  (`CellValue[] | { [columnKey: string]: CellValue }` for a row; a sparse `CellValue[]` for a
  column), consistent with the whole-sheet accessor — one union, used everywhere row/column value
  payloads surface, so the three accessors cannot drift apart.
- A scalar type (`string`, `any`) on any `.values` member is a defect: the runtime always yields the
  sparse collection. A type-level test pins each `.values` member to the `RowValues`/sparse-array
  shape so a future refactor cannot silently narrow it to a scalar again.

### Open questions

- Should the sparse/1-indexed nature be encoded more strongly (e.g. a branded `SparseArray<T>`), or just documented? A branded type is more honest but heavier for consumers; document the 1-indexing at minimum.
- When a fresh, modern API is designed for this fork, consider whether a dense, 0-indexed structure or an explicit `{ rowNumber, values }[]` is a better shape than perpetuating the legacy sparse array.

Related: `public-type-surface-matches-runtime`, `column-level-value-type`.
