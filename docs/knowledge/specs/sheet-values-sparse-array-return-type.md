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

### Open questions

- Should the sparse/1-indexed nature be encoded more strongly (e.g. a branded `SparseArray<T>`), or just documented? A branded type is more honest but heavier for consumers; document the 1-indexing at minimum.
- When a fresh, modern API is designed for this fork, consider whether a dense, 0-indexed structure or an explicit `{ rowNumber, values }[]` is a better shape than perpetuating the legacy sparse array.
