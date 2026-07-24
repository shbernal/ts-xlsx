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

## Test topology

Tests live in two places on purpose, because they are two different kinds of test with
opposite contracts. Where a test goes is decided by *what it is allowed to know*, not by
tidiness:

- **Co-located unit tests — `src/**/*.test.ts`.** White-box. Each sits next to the module
  it exercises (`address.ts` ↔ `address.test.ts`), imports src internals freely, and moves
  or dies with that module under refactor. Co-location keeps the test honest about one unit
  and makes an untested module visible at a glance. Run by `test:src`.
- **The regression corpus — `test/corpus/`.** Black-box and **implementation-blind** (see
  above): cases reach the implementation *only* through the adapter and must never import a
  src internal, because that blindness is the whole reason the corpus outlived the rewrite.
  It is a behavioral spec, not a test of any module. Run by `corpus`.
- **External-oracle harness — `test/ooxml-validation/`.** Validates emitted packages
  against the independent `OpenXmlValidator` (ADR-0002); different toolchain (dotnet),
  different cadence (`test:ooxml`, not in the default `test`).

The wall matters: the `test/` trees earn their separation by being forbidden from reaching
into src the way a co-located unit test may. Put a white-box test in `src/`; keep `test/`
for the blind corpus and the external oracles. A "corpus" case that imports a src internal
has quietly stopped being implementation-blind — the directory boundary is what keeps that
mistake hard to make by accident.

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
| streaming | bounded-memory row streaming — reads, and an incremental workbook writer |
| csv | a thin, optional entry point, never coupled to the xlsx core |
| vba | native read/author/edit of a macro-enabled workbook's `vbaProject.bin` (`src/vba/`) |

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

## The VBA subsystem

Macro-enabled workbooks (`.xlsm`/`.xltm`) carry their VBA as a single opaque part,
`vbaProject.bin` — an OLE2 / Compound File ([MS-CFB]) container holding Office
run-length-compressed ([MS-OVBA]) module source and p-code. OOXML treats it as a binary
blob referenced by a workbook relationship; the ZIP package is otherwise identical to a plain
`.xlsx`. `src/vba/` is the self-contained, dependency-free subsystem that reads that blob and
applies pure-TS structural edits natively — no `fflate`, no runtime dependency, just bytes.
Authoring or editing module *source* is not here: it needs genuinely compiled p-code only a real
Excel can emit, so it lives in the offline `tools/vba-compiler` (see below, and ADR 0019).

**Preservation is the safety floor, and every VBA feature is additive over it.** Read captures
`vbaProject.bin` (and its relationship/content-type closure, including a sibling
`vbaProjectSignature`) as a `PreservedWorkbookReference`, and the writer re-emits those bytes
verbatim — so a load/edit/save of an `.xlsm` keeps its macros with no VBA-specific code on the
common path. Everything below layers onto that guarantee; none of it can desync the two
representations, because a *read-and-not-re-authored* project is still emitted from the
preserved bytes alone.

The subsystem is built as encode/decode pairs over two formats plus a project layer:

- **CFB container** — `cfb.ts` (reader) and `cfb-writer.ts` (writer). The reader walks the
  header → FAT → directory (+ mini-FAT) and reconstructs the whole storage/stream tree so the
  edit path can re-emit it with one stream swapped; the writer emits a v3 container whose
  storages are the name-ordered balanced red-black tree a *navigating* host (Excel) needs, not
  just the linear scan our own reader would accept.
- **MS-OVBA compression** — `ms-ovba.ts` (`decompressContainer` / `compressContainer`): the
  chunked copy-token/literal-run codec Office uses for module source and the `dir` stream (it
  is *not* deflate). The compressor's contract is that its output re-expands byte-for-byte.
- **Project layer** — `project.ts` decodes the `VBA/dir` and module streams into a typed view;
  `project-editor.ts` splices structural edits (remove module, add reference) into an existing
  project; `vba-encoding.ts` holds the shared `dir`-record TLV encoders and VBA name validation.
  `codepage.ts` handles the project code page (MBCS, not latin1) in both directions, and
  `errors.ts` holds the two failure types. (There is deliberately no from-source synthesizer —
  see below.)

**The reader is hostile-input-facing (CLAUDE.md §3).** Every CFB sector index, chain, and
stream size is bounds-checked and cycle-guarded; every MS-OVBA back-reference is validated and
total output is bomb-capped. A malformed project fails closed with `VbaParseError` — never a
crash, hang, or unbounded allocation — each guard pinned by a crafted-malformed fixture. The
authoring/encode side is *our own* bytes, so it fails closed with `VbaAuthorError` on a
contract violation (over-long or duplicate stream name, unrepresentable character) rather than
emitting a silently broken container.

The public surface layers by fidelity and intent, each slice fail-closed:

- **Read** — `Workbook.vbaProject: VbaProject | undefined` parses the preserved bytes *lazily*
  and memoises; modules expose `name`, `streamName`, `kind` (the full procedural / document /
  class / designer classification), and decompressed `source`. A read never perturbs what the
  writer emits (ADR 0016).
- **Attach / replace / strip** — `Workbook.vbaProjectBytes` is a get/set accessor pair over the
  raw blob; a set is validated by `parseVbaProject` *before* any state change, so a bad blob is
  rejected whole, and replacing or removing drops the old bytes' now-stale signature. This is also
  how an authored project is installed: the offline `tools/vba-compiler` produces a compiled
  `vbaProject.bin`, and a consumer attaches it here — pure TS, no Office at runtime.
- **Structural edits (pure TS)** — `removeVbaModule` and `addVbaReference` (with their `Workbook`
  and package-level `editXlsxVba*` wrappers) splice the original `.bin`: they edit only the `dir`
  stream (and, for a removal, `PROJECT`/`PROJECTwm`) and leave every module stream and
  `_VBA_PROJECT` byte-for-byte. They are safe *precisely because* they never touch a module's
  compiled p-code — the `dir` stream, authoritative for the module/reference list, carries the
  change (ADR 0018/0019).
- **Author / edit module source (offline, not in the shipped library)** — done by
  `tools/vba-compiler`, which drives a real headless Excel through the VBIDE object model to emit
  genuinely compiled, source-matched p-code (a `vbaProject.bin`, or a whole edited `.xlsm` for
  document code-behind). It is a Windows+Excel build tool, never in CI, whose output seeds
  committed fixtures (ADR 0019).

**Why source authoring cannot be pure TS.** Excel does **not** recompile VBA from source on open —
a module runs the compiled p-code (PerformanceCache) it ships, and the source is only recompiled
when a human opens the VBE. A `vbaProject.bin` synthesized from source alone (no/mismatched p-code)
either throws "Invalid data format" or *silently runs stale code* — "opens clean" is not proof of
correctness. Only a real Excel can produce runnable p-code; hence the offline compiler. Authoring
artifacts are verified with `execute-verdict.ps1` (opens with macros enabled and *runs* a known
authored macro), not merely by an open verdict (ADR 0019). *Executing* macros in-process is still
out of scope forever — that needs a live host; this is a document tool, not a VBA interpreter
(ADR 0013).

## Tech decisions

The stack is deliberately small and each choice is recorded as an ADR under
[`docs/decisions/`](decisions/):

- **Runtime & no-build dev path** — ADR-0001.
- **Toolchain** (Biome for lint/format; `node --test` over Vitest; hand-rolled
  type-level tests) — ADR-0002. The harness (`test/` + `scripts/`) is TypeScript held
  to the same strict bar as `src/`, gated by `typecheck:test` — ADR-0011.
- **Zip & XML write path** (`fflate`; a hand-written SAX reader with bounded allocation
  on every parser path) — ADR-0003.
- **Docs generated from the types** — ADR-0006.
- **Spec reference** (vendored OOXML schemas + Microsoft Learn MCP) — ADR-0007.
- **VBA subsystem** (read view — ADR-0016; pure-TS structural edits — ADR-0018/0019; source
  authoring moved to the offline `tools/vba-compiler` after the "recompile cookie" premise was
  retracted — ADR-0019; ADRs 0017/0018's from-source mechanism is retracted). Excel as a test
  oracle — ADR-0013.

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
