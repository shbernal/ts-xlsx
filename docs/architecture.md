# Architecture

How `ts-xlsx` is built and why. `CLAUDE.md` is the constitution (the principles every
change answers to); this document describes the library as it stands and the working
agreements that keep it coherent. Point-in-time decisions live under
[`docs/decisions/`](decisions/) as ADRs.

## Origin

`ts-xlsx` is a hard fork of [ExcelJS](https://github.com/exceljs/exceljs), which had
gone effectively unmaintained while still serving tens of millions of downloads a month.
The two assets trapped in that project were of opposite kinds: **knowledge** — thousands
of hours of hard-won understanding of how real-world `.xlsx` files behave, scattered
across hundreds of issues and PRs — and **code**, a weakly-typed, callback-flavored tree
with a rotting dependency graph. The strategy was to separate them: harvest the knowledge
into a durable, implementation-blind form, then rebuild the code from scratch against it.

That harvest is complete. Every credible bug, reproduction, and edge case became a
[regression corpus](../test/corpus/) case; the legacy tree is deleted; the runtime
dependency is now [`fflate`](https://github.com/101arrowz/fflate) alone. What remains is
a modern, strict-TypeScript library whose correctness is pinned by the corpus it carried
across.

## The corpus is the spine

The [regression corpus](../test/corpus/) is the product's backbone. Each case encodes
"correct behavior" as **implementation-blind** assertions that run against any
implementation through a thin adapter — so a behavior, once captured, can never silently
regress. This is why the corpus outlived the rewrite: it was written against the
*behavior*, not the code, so it validated the new implementation the same way it
indicted the old one.

The rule that follows: **when in doubt, add a case.** A bug without a corpus case is a
bug that will return. A missing feature is best reported as a corpus case so it is fixed
once and never regresses.

## Spec & schema reference

Correctness is defined by an external standard, so the ground truth lives in the repo
next to the code that answers to it:

- [`schemas/ooxml-transitional/`](../schemas/) — the complete ECMA-376 **Transitional**
  XSD set (what Excel actually emits), vendored verbatim for offline, greppable reference
  while implementing. It is *reference*, not a validator — conformance validation stays
  with the independent `OpenXmlValidator` oracle (ADR-0002). Repo-only; never published.
- [`docs/knowledge/specs/`](knowledge/specs/) — hand-authored, implementation-blind
  behavior notes from the harvest.
- **Microsoft Learn MCP** (`.mcp.json`) — grounded search over Microsoft's Open
  Specifications ([MS-XLSX] et al.) for the Excel-specific deltas the standard omits.

See ADR-0007 for why the static standard is vendored while the evolving prose is an MCP.

## Module layout

The source tree under [`src/`](../src/) is strict-TypeScript, ESM-only, and build-free on
the dev/test path (Node runs the `.ts` sources directly via type-stripping; `tsc` is the
type *checker* and, for publishing, the emitter). The domain decomposition, in dependency
order:

| Area | Role |
| --- | --- |
| core model | `Workbook` / `Worksheet` / `Row` / `Cell`, addresses, styles — the in-memory document |
| xlsx read/write | OOXML parse and serialize; the hardest, highest-value surface |
| streaming | bounded-memory row streaming for reads |
| csv | a thin, optional entry point, never coupled to the xlsx core |

Cell formatting is one named tuple, not six loose fields. `CellStyle` in `core/style.ts`
holds the six OOXML direct-format facets (`fill`, `numFmt`, `font`, `border`, `alignment`,
`protection`); every style-bearing shape — a cell model, a column's defaults, a table column,
a differential (conditional) format, a named style — derives from it rather than re-declaring
the tuple. `CELL_STYLE_FACETS`, derived from `Record<keyof CellStyle, true>`, is the single
facet list the copy loops walk, so adding a facet is a one-line change the compiler forces every
consumer to honour. Applying a style splits by target: `applyCellStyle` drives a `Cell`'s
per-property setters, `assignStyleFacets` copies plain records — two helpers because a cell and a
bag of fields have different write surfaces.

The two largest surfaces — the xlsx reader and writer — are each a **cluster**, not a
monolith, split along the OOXML package's own seams so a change touches one part:

- **read** (`src/io/xlsx/`): `read-opc.ts` (the OPC/relationship layer), `read-styles.ts`
  (`styles.xml`), `read-worksheet.ts` (one sheet), with `read.ts` keeping `readXlsx` and
  the workbook-level wiring. `rich-runs.ts` owns the `<r>`/`<rPr>`/`<t>` run accumulator
  the worksheet and shared-strings parsers share; `cell-accumulator.ts` owns the per-cell
  gathering state machine the buffered and streaming readers both drive (ADR-0004).
- **write** (`src/io/xlsx/`): `package-plan.ts` (the part-graph plan layer), `workbook-xml.ts`
  and `worksheet-xml.ts` (the serialisers), `part-paths.ts` and `relationships.ts` (shared
  OPC primitives), with `write.ts` keeping `writeXlsx` and the `buildPackageParts`
  orchestrator.

Namespace URIs and ext-URI GUIDs are registered once in `namespaces.ts`. Sheet-local
relationship ids are handed out by a single monotonic `SheetRelIds` allocator: id prefixes
were once re-derived by hand-summing every prior part's count, which silently collides two
parts onto one id when a prefix drifts — ids are now unique by construction and never
recomputed by arithmetic.

The public surface is a single curated barrel, [`src/index.ts`](../src/index.ts). It is
curated, not exhaustive: modelled core-feature types (autofilter, page setup, sheet views,
defined names, image options) are public, and internal helper functions stay off the barrel.
The two halves of the streaming API are symmetric in reach but asymmetric in what needed
naming: the streaming *writer*'s whole surface is public (its workbook/worksheet/row handles
are classes, its options are interfaces — nothing structural is left un-named), while the
streaming *reader*'s per-row/cell output stays inferred-structural rather than a named
commitment. The [API reference](api/README.md) is generated straight from the barrel, so it
cannot describe a shape the compiler wouldn't accept.

## Tech decisions

The stack is deliberately small and each choice is recorded as an ADR under
[`docs/decisions/`](decisions/):

- **Runtime & no-build dev path** — ADR-0001.
- **Toolchain** (Biome for lint/format; `node --test` over Vitest; hand-rolled
  type-level tests) — ADR-0002.
- **Zip & XML write path** (`fflate`; a hand-written SAX reader with bounded allocation
  on every parser path) — ADR-0003.
- **Docs generated from the types** — ADR-0006.
- **Spec reference** (vendored OOXML schemas + Microsoft Learn MCP) — ADR-0007.

## Working agreements

- **Preserve provenance as knowledge, not as a link.** Capture the real-world scenario a
  bug taught us — that survives; an upstream issue number does not. Durable artifacts
  (corpus cases, spec notes, commit messages) never cite upstream numbers; the commit
  that lands a change is its account of record.
- **Security- and correctness-first.** Every parser path is hostile-input-facing: no
  unbounded allocation, no zip-bomb naïveté. Entities are decoded but never expanded;
  inflation is bounded by a running output counter, not any declared size.
- **Narrow foreign tokens; never trust them into the model.** An enumerated attribute
  read from a file is admitted only through a type guard that recognises the known union
  members (the pattern is `isCustomFilterOperator` in `core/autofilter.ts`); an
  unrecognised token is *dropped* — the facet stays unset — rather than cast in with `as`.
  The reader's posture is *skip, never guess*: a malformed token yields absence, not a
  bogus value the rest of the code will trust. See ADR-0004 for the read path this serves.
- **No half-migrations on main.** Each change is fully green — typed, linted, tested,
  corpus-passing — and leaves the tree better than it found it.
