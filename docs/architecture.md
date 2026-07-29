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
| errors | the failure taxonomy every layer throws through (`src/errors.ts`) — below all of them |
| core model | `Workbook` / `Worksheet` / `Row` / `Column` / `Cell`, addresses, styles — the in-memory document |
| xml | escaping, emission and a hostile-input-safe SAX reader (`src/xml/`) — no spreadsheet knowledge |
| opc container | ZIP inflation under a bound, magic-byte sniffing, the relationship graph and part paths (`src/io/opc/`) |
| resolved format | `XfStyle` and what applying an xf to a cell means, shared by both codecs (`src/io/style/`) |
| xlsx read/write | OOXML parse and serialize; the hardest, highest-value surface |
| xlsb read | the binary BIFF12 serialisation of the same model, read-only so far (`src/io/xlsb/`) |
| streaming | bounded-memory row streaming — reads, and an incremental workbook writer |
| csv | a thin, optional entry point, never coupled to the xlsx core |
| vba | native read/author/edit of a macro-enabled workbook's `vbaProject.bin` (`src/vba/`) |

That order is a real constraint, not a description: `scripts/check-layering.ts` (a gate in
`verify --full`) fails the build on an import that runs up it. The rules it carries are that
`src/errors.ts` reaches nothing, `src/xml/` reaches nothing above it, `src/core/` never reaches a
serialisation, `src/io/opc/` and
`src/io/style/` sit below every codec, and the two codecs are peers — `src/io/xlsb/` may not
import `src/io/xlsx/`. Co-located tests are exempt, since a test import is not a dependency of
the graph we ship. Shared code that tempts a codec to reach sideways belongs in `opc` or `style`;
that is what those directories are for.

Cell formatting is one named tuple, not six loose fields. `CellStyle` in `core/style.ts`
holds the six OOXML direct-format facets (`fill`, `numFmt`, `font`, `border`, `alignment`,
`protection`); every style-bearing shape — a cell model, a column's defaults, a table column,
a differential (conditional) format, a named style — derives from it rather than re-declaring
the tuple. `CELL_STYLE_FACETS`, derived from `Record<keyof CellStyle, true>`, is the single
facet list the copy loops walk, so adding a facet is a one-line change the compiler forces every
consumer to honour. Applying a style splits by target: `applyCellStyle` drives a `Cell`'s
per-property setters, `assignStyleFacets` copies plain records — two helpers because a cell and a
bag of fields have different write surfaces.

The same pattern governs a sheet's snapshot one level up. `WorksheetModel` is what
`dst.model = src.model` carries, and a field the getter emits but the setter ignores loses data
silently — the merge-loss failure that contract exists to prevent. Both directions are therefore
driven from one table, `WORKSHEET_MODEL_FACETS` in `core/worksheet-model.ts`, where each field
declares its read and its write side by side along with the clone strategy that field needs
(`{...spread}`, `replaceContents`, replay through the authoring API). Declaration order is the
order a model assignment applies: cells are placed before any merge exists, so a covered cell's
value lands where the model says instead of being routed to a region master mid-load. The registry
is proved exhaustive over `keyof WorksheetModel` at compile time, so a field added without a facet
is a build error that names the field.

Not everything the codecs need to do to the model belongs in its public API. Pushing preserved bytes
back into a `Workbook`, restoring a loaded sheet's hashed protection credential, placing a cell at an
exact position without resolving merges, evicting a row the streaming writer has already serialised —
about fifteen operations exist for a codec's benefit and put the model in states no authoring path can
produce. They were public members: they shipped in the `.d.ts`, they appeared in the generated
reference, and the model class *was* the codec's mutation interface. They now hang off one symbol key
in `core/internal.ts` (`sheet[INTERNAL].restoreProtection(…)`), which no entry barrel exports, so the
boundary is the module graph rather than a naming convention. No layering rule guards who may import
that module: `package.json` maps only the seven subpaths, so the symbol is already unreachable from
outside the package, and a rule would only police `src/core` against itself.

The channel takes two shapes on purpose. `Workbook` and `Worksheet` carry a symbol-keyed *object* of
operations — one allocation per book or sheet, which is nothing. `Cell` gets a symbol-keyed accessor
*pair* instead, because a per-instance channel object on the one class allocated in the millions is a
real cost for state most cells never carry. Despite the origin it is not the *codec* channel: the
model's own `dst.model = src.model` setter reaches through it for exact-position cell placement, and
`Row`/`Column` reach through it for the per-line stores. It is the internal channel, and core is
allowed to use it.

`Row` and `Column` (`core/row.ts`, `core/column.ts`) are **handles, not records**. `Worksheet`
keeps the authoritative stores — the row-major cell grid and the two sparse maps of per-line
formatting — and a handle holds nothing but the sheet and a position, reading and writing straight
through. That is deliberate: a row object that copied its cells out would be the shape of the
merge-loss bug the model contract exists to prevent, and two handles on the same number could
disagree. Position is fixed at construction, the rule `Cell` already follows — a splice that moves
content past `sheet.getRow(3)` does not carry the handle along, any more than it re-points a `Cell`.
Formatting is created on write and never on read, so `getRow(500)` costs nothing and does not extend
the used range; the format record appears when a value is set. The handles reach the stores through
`sheet[INTERNAL]`, which is the same channel the codecs use.

Each handle mirrors its record's fields as accessors — five for a row, eleven for a column (the
geometry plus the six inherited `CellStyle` facets) — so `row.height = 20` is the flat, discoverable
path rather than a hop through a properties bag. That mirror is proved complete at compile time the
same way the model registry is: a field added to `RowProperties` or `ColumnProperties` with no
accessor fails the build naming the field, because otherwise the record would carry it, the codecs
would read and write it, and the public handle would simply never mention it.

The two largest surfaces — the xlsx reader and writer — are each a **cluster**, not a
monolith, split along the OOXML package's own seams so a change touches one part:

- **read** (`src/io/xlsx/`): `read-styles.ts` (`styles.xml`), `read-worksheet.ts` (one sheet),
  with `read.ts` keeping `readXlsx` and the workbook-level wiring. `rich-runs.ts` owns the
  `<r>`/`<rPr>`/`<t>` run accumulator the worksheet and shared-strings parsers share;
  `cell-accumulator.ts` owns the per-cell gathering state machine the buffered and streaming
  readers both drive (ADR-0004).
- **write** (`src/io/xlsx/`): `package-plan.ts` (the part-graph plan layer), `workbook-xml.ts`
  and `worksheet-xml.ts` (the serialisers), `relationships.ts` (the SpreadsheetML relationship-type
  vocabulary), with `write.ts` keeping `writeXlsx` and the `buildPackageParts` orchestrator.

Neither cluster owns the container it rides in. `src/io/opc/` holds what is true of *any* OOXML
package — `inflate.ts` (the bounded inflater), `sniff-format.ts` (the magic-byte probe and the
typed rejection), `read-opc.ts` (resolving relationships and walking a part closure), `rels.ts`
(emitting a `.rels` part), `part-paths.ts`, and the package namespaces — and `src/xml/` holds
escaping, emission and the SAX reader beneath even that.

## How a failure is reported

Every error the library raises deliberately descends from `XlsxError` (`src/errors.ts`), so one
`catch` clause answers "was that us?" without naming a class. Two levels of branch sit under it,
chosen so neither is redundant with the other: `code` is the *kind* of failure, `name` (and
`instanceof`) is exactly which one. Several classes share a code on purpose — a code in 1:1
correspondence with the classes would carry nothing the class did not already carry.

| code | what the caller does about it | classes |
| --- | --- | --- |
| `unsupported-format` | try a different reader, or reject the input | `UnsupportedFormatError` (its `format` field says *which* unsupported input) |
| `malformed-input` | reject the file — it is broken, or hostile | `PackageReadError`, `XmlParseError`, `XlsxParseError`, `XlsbParseError`, `VbaParseError`, `CustomUiParseError` |
| `authoring` | fix the calling code — it described a document that cannot exist | `AuthoringError`, `VbaAuthorError` |
| `internal` | report it — an invariant of ours did not hold | `InternalError` |

Scalar argument validation stays outside the taxonomy: an index out of range, an unparseable
reference, a value of the wrong type are native `RangeError` / `SyntaxError` / `TypeError`, because
that is what those types are for. The line is composite vs. scalar — `getColumn(0)` is a
`RangeError`; a table that names a column twice is an `AuthoringError`. A layer that wraps a
lower-level failure passes it as `cause` rather than flattening it into the message, except where
the lower layer's text is itself the hazard (a zip library's message can name an absolute path, so
`sniff-format.ts` replaces rather than wraps it).

There is deliberately no "not implemented yet" code: every candidate turned out to be an
unreachable exhaustiveness guard, and the one real feature gap — a binary `.xlsb` cannot be
row-streamed — is already reported through `UnsupportedFormatError`'s `format` branch.

## Two serialisations, one model

`.xlsb` is not a second library bolted on; it is a second **codec** over the same `Workbook`. The two
formats share an OPC/ZIP container, a relationship graph, and a style model, and differ only in how the
office-document parts are spelled — XML in `.xlsx`, BIFF12 record streams in `.xlsb`. The code follows
that seam exactly, and the directory layout states it: the bounded inflater, magic-byte probe and
OPC/relationship resolution live in `src/io/opc/`, the resolved-format table (`XfStyle`, its built-in
number formats, and `applyXfToCell`) in `src/io/style/` — both *above* the codecs rather than inside
either — and only the part parsers
live apart in `src/io/xlsb/` — `record-stream.ts` (the record framing), `primitives.ts` (RkNumber,
length-prefixed strings, colours), `formula.ts` and `ptg-functions.ts` (the Ptg token stream a binary
formula is stored as, decoded to the text `<f>` would have carried), then a per-part parser mirroring
its XML counterpart.

`readXlsx` detects which serialisation a package holds and dispatches, so a caller never branches on
format. The property that keeps the two honest is asserted, not assumed: the corpus reads one workbook
Excel saved in *both* forms and requires the two models to be identical. Anything the binary states
that XML omits — a bottom vertical alignment, a locked cell, a row restating the sheet's default
height, a hatch fill's automatic-colour sentinels — must therefore be dropped on the binary side, which
is where most of the subtlety in that reader lives.

Namespace URIs and ext-URI GUIDs are registered once, split by which layer owns them: the
package-level URIs (`.rels`, content types, the relationship vocabulary) in
`src/io/opc/namespaces.ts`, the SpreadsheetML and extension ones in `src/io/xlsx/namespaces.ts`.
Sheet-local
relationship ids are handed out by a single monotonic `SheetRelIds` allocator: id prefixes
were once re-derived by hand-summing every prior part's count, which silently collides two
parts onto one id when a prefix drifts — ids are now unique by construction and never
recomputed by arithmetic.

The public surface is seven curated entry barrels under [`src/entries/`](../src/entries/), one
per subpath the package publishes — `/core`, `/xlsx`, `/xlsb`, `/csv`, `/vba`, `/customui`,
`/errors` — and [`src/index.ts`](../src/index.ts), which unions them so the bare package name
still carries everything. Each symbol is listed in exactly one entry, so the root barrel is a
union of `export *` lines rather than a second list to keep in step.

That disjointness is load-bearing rather than tidy: `export *` does not report an ambiguous
re-export, it silently drops the name, so a symbol exported from two entries would vanish from
the root specifier with no diagnostic anywhere. `scripts/check-entries.ts` (a gate in
`verify --full`) fails the build on a duplicate, on an entry `package.json` does not publish, and
on a published subpath whose module is gone. It is also why the *whole* failure taxonomy is
exported from `/errors` and nowhere else: a container-level failure belongs to no single codec —
`readXlsx` and `readXlsb` both raise `UnsupportedFormatError` — so putting the classes with the
codecs would have forced exactly the duplication the union cannot survive. That entry costs 12 KB,
so classifying a failure never loads a parser.

The barrels are curated, not exhaustive: modelled core-feature types (autofilter, page setup,
sheet views, defined names, image options) are public, and internal helper functions stay off
them. `src/vba/index.ts` and `src/customui/index.ts` are *internal* barrels that the model and
the codecs import; the public `/vba` and `/customui` faces are deliberately narrower. The two
halves of the streaming API are symmetric in reach but asymmetric in what needed naming: the
streaming *writer*'s whole surface is public (its workbook/worksheet/row handles are classes, its
options are interfaces — nothing structural is left un-named), while the streaming *reader*'s
per-row/cell output stays inferred-structural rather than a named commitment. Streaming is not
its own subpath: measured, it reaches every module `/xlsx` does plus three, and an entry that
costs what the codec costs is an alias rather than a boundary.

`"sideEffects": false` is declared, and it is true — no module in `src/` mutates anything at
import time, so a bundler may drop an unused one whole. `scripts/size-budget.ts` measures each
subpath's static-import closure against its own budget, which is what notices a boundary being
crossed: a codec acquiring a value-import of something it previously needed only as a type moves
one of those numbers and leaves the package total untouched. `scripts/smoke-dist.ts` resolves
every subpath through the package name — self-reference, the only thing in the repo that
exercises the `exports` map at all — and asserts `/core` reaches no serialisation and `/errors`
reaches nothing but itself.

Those numbers say one thing about the model that is worth stating rather than leaving to be
rediscovered: roughly half of `/core`'s weight is the VBA codec and the ribbon parser, because
`core/workbook.ts` imports `parseVbaProject`, `addVbaReference` and `removeVbaModule` at *value*
level. The layering rule permits that deliberately — a workbook models a VBA project and its ribbon,
so those types are part of the document — and the measurement is what shows the price of letting the
operations live there too. It is left alone knowingly. Removing it means a seam the model does not
import at value level, which breaks `Workbook`'s VBA methods, and it would pay off only for a
consumer importing `/core` with no bundler at all; with a bundler, `sideEffects: false` already
prunes what such a consumer never calls. Revisit it when that consumer turns out to exist rather
than on the strength of the number alone — and until then `./core`'s budget is what notices the
weight growing further.

The [API reference](api/README.md) is generated straight from the root barrel, so it cannot
describe a shape the compiler wouldn't accept.

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
  to the same strict bar as `src/`, gated by `typecheck:test` — ADR-0011. All of those
  gates run as one concurrent, content-cached command (`node scripts/verify.ts`) —
  ADR-0022.
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
- **On a round-tripping surface, ask whether input is safe to *write back*.** Bounding what a
  parser will hold is only half of it: what the reader accepts, the writer re-emits, so a value
  a foreign part carries can leave *our* output invalid — and one invalid attribute is enough for
  Excel to offer to repair the feature away. A wire bound therefore lives in the parser, the
  authoring verb, *and* the serialiser, so the serialiser cannot emit an illegal value however the
  model was populated (`MENTION_OFFSET_MAX` is the worked example; see
  `knowledge/specs/threaded-comments-and-the-legacy-fallback.md`).
- **Two parts that are two halves of one representation derive from one variable.** Some features
  are only coherent as a pair — a threaded comment's conversation part and the legacy fallback
  `<comment>` that binds a cell to it are each *invisible in Excel* without the other, even though
  either alone validates clean. The writer computes such a pair from a single source (`write.ts`
  derives both from one `threads` value) so neither can be emitted without the other by
  construction, rather than by a rule someone has to remember.
- **A part family graduates from preserved to modeled in one change, never both at once.**
  Byte preservation (`core/preserved.ts`) is the sole emission authority for what it covers, so a
  *read view* over preserved bytes is safely additive (`Workbook.vbaProject`, `Workbook.customUI`,
  `loadedPivotTables`) — but the moment a serialiser emits a part from the model, that rel type
  must leave `isPreservedSheetRelType`/`isPreservedWorkbookRelType` in the same commit or the
  package carries it twice.
- **Narrow foreign tokens; never trust them into the model.** An enumerated attribute
  read from a file is admitted only through a type guard that recognises the known union
  members (the pattern is `isCustomFilterOperator` in `core/autofilter.ts`); an
  unrecognised token is *dropped* — the facet stays unset — rather than cast in with `as`.
  The reader's posture is *skip, never guess*: a malformed token yields absence, not a
  bogus value the rest of the code will trust. See ADR-0004 for the read path this serves.
- **No half-migrations on main.** Each change is fully green — typed, linted, tested,
  corpus-passing — and leaves the tree better than it found it.
