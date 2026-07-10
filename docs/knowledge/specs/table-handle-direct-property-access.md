# A table accessor must return a usable table handle, not a nested wrapper

Cluster: tables

## Scenario

A user adds a table to a worksheet and then wants to read or modify its metadata — the columns, the
anchoring reference range, the rows, the applied style — and to rename a column and save. The
natural, discoverable API is for the add-table and get-table accessors to return a table **handle**
whose documented properties are directly readable and writable. In the reference implementation these
accessors instead returned an internal wrapper of the shape `{ worksheet, table }`: the typed
properties (`columns`, `ref`, `rows`, `style`) were `undefined` at the top level, and the real data
lived under `handle.table.*`. The types were the docs, and the docs lied. Worse, a column rename
applied to the wrong (wrapper) surface silently produced a **corrupt file** rather than erroring.

> Spec note, not a corpus case: this is an API-shape and type-honesty redesign, not a single failing
> behavior. The durable value is the desired handle contract and the naming/mutation decisions.

## Desired behavior

- The create/retrieve accessors return **one** table handle object — not a wrapper bundling the
  worksheet alongside a nested table.
- The handle's documented properties — column definitions, anchoring reference range, rows, applied
  style/theme — are directly accessible and consistent with the declared types. A read-after-write
  of a freshly added table returns the values that were supplied.
- Mutating table metadata through the handle — in particular **renaming a column** — is reflected on
  save and never corrupts the package: the file opens cleanly and round-trips.
- `add`, `get`, and `remove` share **one table-identity concept** (by name) so their
  return/argument types line up as a coherent trio.

## Open questions

- The reference-range property name: the reference implementation used both `ref` (interface) and
  `tableRef` (runtime); these collapse to one honest name.
- Live view vs snapshot: should the handle be a live view (mutations flow into the workbook model)
  or an immutable snapshot with explicit mutation methods? Given the fork's immutable-by-default
  stance, prefer explicit typed mutation entry points over free property assignment — while the
  **read** surface still exposes `columns`/`ref`/`rows`/`style` directly.

Related: `existing-table-roundtrip-fidelity`, `loaded-table-exposes-data-rows`,
`table-headerless-omits-autofilter`.
