# The public type surface must expose every real runtime accessor and option

Cluster: types

## Scenario

The legacy library shipped hand-authored `.d.ts` declarations that drifted from the runtime: real
accessors and options existed at runtime but were absent from the published types, forcing consumers
into `any` casts or module augmentation to use them. Two concrete instances from the backlog:

1. **Worksheet data-validation management** — the runtime supports adding, finding, and removing
   per-range data validations, but the published `Worksheet` type omits the accessor, so a
   TypeScript caller attaching a dropdown to a range cannot do so without casting.
2. **Worksheet protection `spinCount`** — the OOXML `sheetProtection` element supports a `spinCount`
   attribute (the password-hash iteration work factor), and the runtime honors it, but the
   `WorksheetProtection` options type omits it, so a caller cannot configure the hardening factor in
   a typed way.

> Spec note, not a corpus case: the runtime behavior largely exists — the defect is a type-surface
> completeness gap, pinned by type-level tests (`expectTypeOf`/tsd) plus a behavioral round-trip
> assertion, not by a data-file corpus case. The durable value is the principle and the specific
> members that must be typed.

## Desired behavior

- **The public, strictly-typed surface exposes every real runtime capability** — no `any` casts, no
  consumer-side module augmentation required. For a TypeScript-first library the types ARE the docs;
  a runtime accessor missing from the types is a defect.
- **Data validations are first-class and precisely typed**: add a validation to an address/range with
  a fully-typed descriptor, look up the validation for an address (descriptor or a clear absence
  value), and remove it. The descriptor is a **discriminated union keyed on kind** (`list`, `whole`,
  `decimal`, `date`, `textLength`, `custom`, …) so only the fields meaningful for that kind are
  present.
- **Worksheet protection accepts `spinCount`** alongside the password and flags; on write it is
  emitted on `sheetProtection` (with the algorithm/hashValue/saltValue from password hashing) and on
  read it round-trips, so the configured work factor is preserved. The type includes it by
  construction.
- **Guarded by construction**: because the types are generated from the source rather than hand-
  authored, a runtime member cannot silently be missing from the declarations — and a CI type-level
  test asserts the presence and precision of the public surface.

## Open questions

- The accessor naming and shape for data validations (a `dataValidations` manager object vs
  per-cell/`getCell(addr).dataValidation`), and how the discriminated descriptor is exported.
- Default `spinCount` (Excel commonly emits 100000) and whether it is capped/bounded on read to
  avoid a hostile file forcing an extreme iteration count (ties to the encryption note's bounding).
- How broadly to audit the rest of the surface for the same drift, and whether a generated-types
  pipeline plus tsd tests fully closes the class of bug.

Related: `published-types-resolve-across-consumers`, `public-types-node-stream-portability`,
`write-buffer-return-type-contract`, `sheet-protection-permits-requested-operations`,
`workbook-password-encryption`.
