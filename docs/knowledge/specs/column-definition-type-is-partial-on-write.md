# Column definitions must be a partial write-side type, distinct from the read-side column

Cluster: types

## Scenario

A TypeScript user assigns an array of column definitions to a worksheet's `columns`, supplying only
the fields they care about — a header, a key, a width — and expects it to type-check. In practice
the write-side column type demanded every property of the full column model (`outlineLevel`,
`hidden`, `style`, `values`, and a dozen more), so a minimal object literal was rejected by the
compiler even though it is valid input at runtime. Users pinned an older version or cast their
literals to silence the error. The tension: the type describing a column **read back** (a
fully-populated, computed object with helper methods) is not the shape a caller should have to supply
when **defining** columns — on write, only `header`, `key`, `width`, `outlineLevel`, `hidden`, and
`style` are meaningful, and all should be optional.

## Desired behavior

- The public API distinguishes the **write-side column-definition type** from the **read-side column
  type**. When a caller assigns column definitions, every field is optional and only author-supplied
  properties are accepted (`header`, `key`, `width`, `outlineLevel`, `hidden`, `style`). A minimal
  literal such as `{ header, key, width }` type-checks **without a cast**.
- The read-side column returned by the library may be a richer object (fully populated, with
  computed/derived properties and helper methods), but callers are never required to construct that
  shape.
- Lock this as a **type-level test** so a future refactor cannot silently re-require the full shape.

## Prior art

The earlier upstream type modeled the write assignment as an array of partial column objects, which
type-checked correctly; a later change tightened it to require the full column interface, breaking
all minimal definitions, and it was reverted back to partial. That history confirms the correct
model: **input is partial/optional; the full/required shape belongs only to read-back.** A cleaner
design the community proposed: two explicit types — a plain optional `ColumnDefinition` for input and
a richer `ReadonlyColumn` (extending it with `toString`/`equivalentTo`-style helpers) for output —
and optionally a distinct setter so the read property and the write path do not share one over-broad
type.

## Open questions

- Keep `columns` an assignable property that accepts definitions, or move defining columns to an
  explicit setter while the getter returns the read-side type?
- Which fields are legitimately settable per-column at definition time — confirm the set (`header`,
  `key`, `width`, `outlineLevel`, `hidden`, `style`; `header` may be a single string or an array of
  strings for multi-row headers)?
- Mirror the same partial-input principle for rows and cells so the whole authoring surface is
  consistently minimal-input?

Related: `worksheet-columns-mutable-array-ergonomics`, `column-level-value-type`,
`column-key-roundtrip-persistence`, `declarative-nested-column-headers`,
`public-types-node-stream-portability`, `async-iterable-types-compile-cleanly`.
