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
3. **Streaming writer `stream` accessor** — the streaming `WorkbookWriter` exposes a `stream` property
   at runtime (the underlying writable the package is emitted to), but the published type omits it, so
   a TypeScript caller inspecting or wiring up the stream must cast.
4. **`Range.forEachAddress` iteration** — the `Range` model exposes an address-iteration callback at
   runtime (walk every cell address in the range), but the published type omits it, forcing `@ts-ignore`
   on the exact ergonomic loop it exists to serve (applying a style to each address in a block).
5. **`Range` is a value, not just a shape** — the runtime exports `Range` as a constructable class,
   but the published types declared only an `interface Range` (a structural shape with no
   constructor). A caller doing `new Range(...)` — the documented way to build one — got a "not a
   constructor" type error and had to deep-import the internal module with `@ts-ignore`. Any runtime
   value export (a class or a namespace object) must have a *value*-level declaration, not merely a
   same-named type, so `new`, `instanceof`, and static members type-check.
6. **Streaming worksheet reader `id` / `name` / `state`** — iterating the streaming `WorkbookReader`
   yields a per-worksheet reader whose `id`, `name`, and `state` (visible/hidden/veryHidden) the
   library populates at runtime, but the published `WorksheetReader` type omits them. The idiomatic
   loop — `for await (const ws of reader) console.log(ws.id, ws.name, ws.state)` — then errors under
   TypeScript, forcing casts on the exact identity fields the reader exists to expose (confirmed at
   runtime: a hidden sheet iterated through the streaming reader reports its `id`, `name`, and
   `state: 'hidden'`). This is the streaming counterpart of the drift above: the streaming reader
   types must expose every field the reader sets — `id`, `name`, and `state` among them.

7. **`WorkbookReader` options argument is optional** — the streaming `WorkbookReader` can be
   constructed with no options, defaulting every setting; the runtime accepts the options argument
   omitted. The published type declared the options argument (or the options-object type) as
   required, so a TypeScript caller who wants the defaults is forced to pass an empty object or a
   filler value. An options argument the runtime treats as optional must be typed optional, with each
   field's default documented on the type.

8. **CSV read/write options types match the documented shape** — the CSV reader/writer options are
   forwarded to the underlying CSV engine, and the shape shown in the docs (delimiter, headers,
   quote/escape, date parsing, …) is what the runtime accepts; but the published options type
   diverged from that shape, so following the documented example produced a TypeScript error and
   callers cast or `@ts-ignore` the exact options the docs told them to pass. The CSV options type
   must accept the documented, runtime-honored shape without casts — the same generated-from-source
   discipline as the rest of the surface, applied to the CSV boundary.

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
- **Runtime value exports are declared as values**: a class or namespace object that exists at
  runtime is exported with a value-level declaration, so `new`, `instanceof`, and static access
  type-check without deep-importing internals. A same-named `interface` that shadows a real class is
  a defect.
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
