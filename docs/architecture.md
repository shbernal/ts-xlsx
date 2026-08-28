# Architecture

How `ts-xlsx` is built and why. `CLAUDE.md` is the constitution (the principles every
change answers to); this document describes the library as it stands and the working
agreements that keep it coherent. Point-in-time decisions live under
[`docs/decisions/`](decisions/) as ADRs.

## Origin

`ts-xlsx` is a hard fork of [ExcelJS](https://github.com/exceljs/exceljs), which had
gone effectively unmaintained while still serving tens of millions of downloads a month.
Two assets were trapped in that project, and they were of opposite kinds. One was
knowledge: thousands of hours of hard-won understanding of how real-world `.xlsx` files
behave, scattered across hundreds of issues and PRs. The other was code, a weakly-typed,
callback-flavored tree with a rotting dependency graph. The strategy was to separate
them, harvest the knowledge into a durable, implementation-blind form, then rebuild the
code from scratch against it.

That harvest is complete. Every credible bug, reproduction, and edge case became a
[regression corpus](../test/corpus/) case; the legacy tree is deleted; the runtime
dependency is now [`fflate`](https://github.com/101arrowz/fflate) alone. What remains is
a modern, strict-TypeScript library whose correctness is pinned by the corpus it carried
across.

## The corpus defines correct behavior

The [regression corpus](../test/corpus/) is what pins this library down. Each case
encodes "correct behavior" as implementation-blind assertions that run against any
implementation through a thin adapter, so a behavior, once captured, can never silently
regress. This is why the corpus outlived the rewrite. It was written against the
*behavior* and not the code, so it validated the new implementation the same way it
indicted the old one.

The rule that follows: when in doubt, add a case. A bug without a corpus case is a bug
that will return. A missing feature is best reported as a corpus case so it is fixed once
and never regresses.

## Test topology

Tests live in two places on purpose, because they are two different kinds of test with
opposite contracts. Where a test goes is decided by *what it is allowed to know*, not by
tidiness:

- **Co-located unit tests, `src/**/*.test.ts`.** White-box. Each sits next to the module
  it exercises (`address.ts` ↔ `address.test.ts`), imports src internals freely, and moves
  or dies with that module under refactor. Co-location keeps the test honest about one unit
  and makes an untested module visible at a glance. Run by `test:src`.
- **The regression corpus, `test/corpus/`.** Black-box and implementation-blind (see
  above): cases reach the implementation *only* through the adapter and must never import a
  src internal, because that blindness is the whole reason the corpus outlived the rewrite.
  It is a behavioral spec, not a test of any module. Run by `corpus`.
- **External oracle tests, `test/ooxml-validation/`.** These validate emitted packages
  against the independent `OpenXmlValidator`, via the shared `ooxml-validate` package
  (ADR-0002). Different toolchain (a .NET binary this repo neither builds nor pins) and
  different cadence (`test:ooxml`, not in the default `test`).

The wall matters: the `test/` trees earn their separation by being forbidden from reaching
into src the way a co-located unit test may. Put a white-box test in `src/`; keep `test/`
for the blind corpus and the external oracles. A "corpus" case that imports a src internal
has quietly stopped being implementation-blind, and the directory boundary is what keeps
that mistake hard to make by accident.

## Spec and schema reference

Correctness is defined by an external standard, so the ground truth lives in the repo
next to the code that answers to it:

- [`.claude/skills/ooxml-lookup/`](../.claude/skills/ooxml-lookup/) holds the ECMA-376
  schema (Transitional and Strict) as a local SQLite graph behind a query CLI, vendored for
  offline, deterministic reference while implementing. Ask it what may go inside an element
  and in what order, what attributes a type takes, and what values those accept, rather
  than hand-joining XSDs:
  `node .claude/skills/ooxml-lookup/scripts/ooxml.mjs children x:c`. It is *reference*,
  not a validator. Conformance validation stays with the independent `OpenXmlValidator`
  oracle (ADR-0002). Repo-only; never published.
- [`docs/knowledge/specs/`](knowledge/specs/) holds hand-authored, implementation-blind
  behavior notes from the harvest.
- Microsoft Learn MCP (`.mcp.json`) is grounded search over Microsoft's Open
  Specifications ([MS-XLSX] et al.) for the Excel-specific deltas the standard omits.

See ADR-0007 for why the static standard is pinned in the tree while the evolving prose is
an MCP, and ADR-0034 for why that pinned form is now a queryable graph rather than the raw
XSD set it used to be.

## Module layout

The source tree under [`src/`](../src/) is strict-TypeScript, ESM-only, and build-free on
the dev/test path (Node runs the `.ts` sources directly via type-stripping; `tsc` is the
type *checker* and, for publishing, the emitter). The domain decomposition, in dependency
order:

| Area | Role |
| --- | --- |
| errors | the failure taxonomy every layer throws through (`src/errors.ts`), below all of them |
| core model | `Workbook` / `Worksheet` / `Row` / `Column` / `Cell`, addresses and styles: the in-memory document |
| xml | escaping, emission and a hostile-input-safe SAX reader (`src/xml/`), with no spreadsheet knowledge |
| opc container | ZIP inflation under a bound, magic-byte sniffing, the relationship graph and part paths (`src/io/opc/`) |
| resolved format | `XfStyle` and what applying an xf to a cell means, shared by both codecs (`src/io/style/`) |
| xlsx read/write | OOXML parse and serialize; the hardest, highest-value code in the tree |
| xlsb read | the binary BIFF12 serialisation of the same model, read-only so far (`src/io/xlsb/`) |
| streaming | bounded-memory row streaming, both reads and an incremental workbook writer |
| csv | a thin, optional entry point, never coupled to the xlsx core |
| vba | native read/author/edit of a macro-enabled workbook's `vbaProject.bin` (`src/vba/`) |

That order is a real constraint, not a description: `scripts/check-layering.ts` (a gate in
`verify --full`) fails the build on an import that runs up it. The rules it carries are that
`src/errors.ts` reaches nothing, `src/xml/` reaches nothing above it, `src/core/` never reaches a
serialisation, `src/io/opc/` and `src/io/style/` sit below every codec, and the two codecs are
peers, so `src/io/xlsb/` may not import `src/io/xlsx/`. Co-located tests are exempt, since a test
import is not a dependency of the graph we ship. Shared code that tempts a codec to reach sideways
belongs in `opc` or `style`; that is what those directories are for.

### The two model classes delegate their state, they do not accumulate it

`Worksheet` and `Workbook` are the two classes everything else hangs off, so both would grow
without bound if every feature simply added a private field beside the getter that reads it. Past
about a thousand lines that is no longer a class you can read: the fields are scattered through the
file, and there is no point at which you can see what the object *is*.

Both push cohesive slices of state into their own objects and keep the public accessors in front of
them. `Worksheet` holds `DataValidationOverlay`, `ConditionalFormattingOverlay`, `GridEdits`,
`WorksheetPictures` (`core/worksheet-pictures.ts`) and `WorksheetComments`
(`core/worksheet-comments.ts`); `Workbook` holds `WorkbookVbaProject` (`core/workbook-vba.ts`),
`WorkbookTheme` (`core/workbook-theme.ts`) and `WorkbookStyleTables` (`core/workbook-styles.ts`).
The public surface does not move: an accessor stays on the model class, keeps its name, its type
and its full doc comment, and becomes a one-line delegation. The doc comment staying put is not
incidental, since it is what `scripts/gen-docs.ts` reads and what a consumer sees; the slice
carries implementation notes only.

What makes a slice a slice is how little it touches outside itself. The VBA project reaches exactly
one thing, the preserved-reference list, which is handed to it. The style tables reach nothing: the
six of them (`<dxfs>`, the named cell styles, the two colour lists, the table-style block and the
definitions a caller authors) are one thing said six ways, a table read out of `styles.xml` and
handed back to the writer, each the target of an index held elsewhere in the file.

The theme reaches one, and it is instructive: resolving a colour needs the workbook's custom indexed
palette as well as the theme scheme, but that palette is also the writer's source for
`<indexedColors>` and is filled in by the reader, so it belongs with the style tables rather than
with the theme. It is passed in as a narrow accessor over that slice. Had the palette moved *into*
the theme, the rest of the styles-table state would have followed it and the result would be a
colour-and-styles overlay, which is not a slice of anything. When a candidate slice has more than one
or two such edges, that is the signal it is not one. The media block on `Workbook` is the standing
example of a candidate that fails it: `exportImages`/`importImages` reach five different things on
`Worksheet`, so grouping them would move the coupling rather than remove it.

`Worksheet` is still over the thousand lines after those two, and deliberately so. What is left on
it is the grid and the things that reach into the grid constantly: tables and pivots materialise
header and totals rows and re-pin themselves through `GridEdits` on every splice, so lifting them
would produce a tables-and-grid overlay, which is the theme example above with a different name.

What that materialising *is*, though, belongs to the table, and lives there. `Worksheet.addTable`
hands the new table a `TableGrid`, the three-call channel a registered table holds into the grid
(does this cell hold a value, write this cell, open a row here), and the table fills its own header
and totals cells as the last act of construction. The knowledge that an empty header row is
corruption Excel repairs on open, that a totals aggregate is a `SUBTOTAL` under a code from
`TOTALS_ROW_SUBTOTAL_CODE`, and that a `custom` total is the column's own stored formula, is the
table's, not the sheet's. The round-trip guard is the load-bearing part: reading a workbook
re-registers every table after its cells are loaded, so only an *empty* cell may be filled. A cell
holding rich text, a style, or text that drifted from the column name is authoritative, and
clobbering it would be silent data loss on every file that has a table, with no schema error to
catch it. That is why the channel asks whether a cell holds a value rather than handing over the
sheet: a materialiser that could reach `#cellAt` would create the very cells it was asking about. The
page-layout fields (`view`, `pageSetup`, `printOptions`, `pageMargins`, `headerFooter`, and the two
break lists) are plain mutable objects with no accessors and no behaviour, so there is nothing to
delegate and grouping them would change the public API to no end. The line count is the symptom the
rule watches for, not the rule; a slice that is not one costs more than the lines it removes.

`GridEdits` owns that splice arithmetic for *everything* anchored to the grid, which is a wider set
than the cell rows: line metadata, merges, tables, anchored images and shared-formula anchors move
with the cells, and so do the four things bound to a range that live outside the cell grid entirely:
data validations, conditional formats, comment threads and the sheet's autofilter. The invariant
is the general one, not a list that happened to be complete once: an overlay left behind re-points a
dropdown or a highlight rule at cells nobody chose, and the writer emits that without complaint. The
single coordinate rule they all share is `core/grid-shift.ts`, and every participant answers the same
two questions through it: where does this land, and did the delete swallow it whole? A range the
delete swallowed whole takes its entry with it rather than clamping onto the cut line, because
dropping a rule is legible and silently re-aiming one is not.

The row and column axes are deliberately not mirror images, and where they diverge is a decision
rather than a gap someone forgot to close. A row takes either input shape, a positional array or an
object keyed by `ColumnProperties.key`; a column takes only the positional one, because a column's
values are indexed by *row* and a row carries no key for the other shape to name. `duplicateRow` has
no column counterpart for a mechanical reason rather than a matter of taste: a row insert carries
pre-built `Cell`s, so a duplicate keeps the source's per-cell styles, while a column insert carries
raw `CellValue`s that materialise fresh cells, so the same verb on that axis would silently drop the
styles it claims to be copying. What is missing is the machinery, not the verb, and adding it means
giving `spliceColumns` a cell-shaped insert path first.

### Inside `src/io/xlsx/`: three kinds of module, deliberately flat

Thirty-odd files in one directory reads like something nobody got round to organising, and the
obvious tidy-up into `read/`, `write/` and `shared/` is wrong here. The modules fall into three
kinds, and the split would cut across the most cohesive of them:

- **Feature modules, both directions in one file.** `comments.ts`, `tables.ts`, `images.ts`,
  `hyperlinks.ts`, `data-validation.ts`, `conditional-formatting.ts`, `threaded-comments.ts` each
  export a `parseX` for the reader beside an `xXml` for the writer. That pairing is the point: the
  two halves share one feature's element names and must agree with each other, and a round-trip is
  exactly the claim that they do. Splitting each into two files would double the count while moving
  the two functions that have to stay in step into different directories.
- **The read pipeline.** `read.ts` and everything prefixed `read-*`, plus its private helpers
  (`cell-accumulator.ts`, `cell-value.ts`, `rich-runs.ts`). The `read-` prefix is the convention;
  `pivot-read.ts` and `shared-strings-read.ts` were the two files spelling it the other way round.
- **The write pipeline.** `write.ts`, `write-stream.ts`, the `*-xml.ts` serialisers, and the
  write-side services (`styles.ts`'s interning registry, `shared-strings.ts`, `package-plan.ts`).

The invariant worth having is that the read pipeline never reaches into the write pipeline. It
holds, and `color-xml.ts` is why it took work: `parseColor` sat in the write-side style *table*, so
the style reader and the worksheet reader both imported the writer to decode a `<color>`. Reading
and writing that element are one concern with two directions, and they now live together in a module
either side may use.

That invariant is documented rather than gated, because a gate for it cannot be honest.
`check-layering.ts` matches directories, and a rule derived from reachability defeats itself: the
moment a read module imports a write module, that module becomes reachable from the read roots and
so classifies as *shared*, which is precisely the label that makes the check pass. A declared list
of write-pipeline modules would work but goes stale in silence, which is the failure these checkers
exist to prevent. See ADR 0030.

Cell formatting is one named tuple, not six loose fields. `CellStyle` in `core/style.ts`
holds the six OOXML direct-format facets (`fill`, `numFmt`, `font`, `border`, `alignment`,
`protection`), and every style-bearing shape derives from it rather than re-declaring the
tuple: a cell model, a column's defaults, a table column, a differential (conditional)
format, a named style. `CELL_STYLE_FACETS`, derived from `Record<keyof CellStyle, true>`, is
the single facet list the copy loops walk, so adding a facet is a one-line change the compiler
forces every consumer to honour. Applying a style splits by target: `applyCellStyle` drives a
`Cell`'s per-property setters, `assignStyleFacets` copies plain records. Two helpers, because a
cell and a bag of fields take writes differently.

The same pattern governs a sheet's snapshot one level up. `WorksheetModel` is what
`dst.model = src.model` carries, and a field the getter emits but the setter ignores loses data
silently, which is the merge-loss failure that contract exists to prevent. Both directions are
therefore driven from one table, `WORKSHEET_MODEL_FACETS` in `core/worksheet-model.ts`, where each
field declares its read and its write side by side along with the clone strategy that field needs
(`{...spread}`, `replaceContents`, replay through the authoring API). Declaration order is the
order a model assignment applies: cells are placed before any merge exists, so a covered cell's
value lands where the model says instead of being routed to a region master mid-load. The registry
is proved exhaustive over `keyof WorksheetModel` at compile time, so a field added without a facet
is a build error that names the field.

Not everything the codecs need to do to the model belongs in its public API. Pushing preserved bytes
back into a `Workbook`, restoring a loaded sheet's hashed protection credential, placing a cell at an
exact position without resolving merges, evicting a row the streaming writer has already serialised:
about fifteen operations exist for a codec's benefit and put the model in states no authoring path
can produce. They were public members. They shipped in the `.d.ts`, they appeared in the generated
reference, and the model class *was* the codec's mutation interface. They now hang off one symbol key
in `core/internal.ts` (`sheet[INTERNAL].restoreProtection(…)`), which no entry barrel exports, so the
boundary is the module graph rather than a naming convention. No layering rule guards who may import
that module: `package.json` maps only the seven subpaths, so the symbol is already unreachable from
outside the package, and a rule would only police `src/core` against itself.

The channel takes two shapes on purpose. `Workbook` and `Worksheet` carry a symbol-keyed *object* of
operations, one allocation per book or sheet, which is nothing. `Cell` gets a symbol-keyed accessor
*pair* instead, because a per-instance channel object on the one class allocated in the millions is a
real cost for state most cells never carry. Despite the origin it is not the *codec* channel: the
model's own `dst.model = src.model` setter reaches through it for exact-position cell placement, and
`Row`/`Column` reach through it for the per-line stores. It is the internal channel, and core is
allowed to use it.

`Row` and `Column` (`core/row.ts`, `core/column.ts`) are handles, not records. `Worksheet`
keeps the authoritative stores, the row-major cell grid and the two sparse maps of per-line
formatting. A handle holds nothing but the sheet and a position, reading and writing straight
through. That is deliberate: a row object that copied its cells out would be the shape of the
merge-loss bug the model contract exists to prevent, and two handles on the same number could
disagree. Position is fixed at construction, the rule `Cell` already follows: a splice that moves
content past `sheet.getRow(3)` does not carry the handle along, any more than it re-points a `Cell`.
Formatting is created on write and never on read, so `getRow(500)` costs nothing and does not extend
the used range; the format record appears when a value is set. The handles reach the stores through
`sheet[INTERNAL]`, which is the same channel the codecs use.

Each handle mirrors its record's fields as accessors, five for a row and eleven for a column (the
geometry plus the six inherited `CellStyle` facets), so `row.height = 20` is the flat, discoverable
path rather than a hop through a properties bag. That mirror is proved complete at compile time the
same way the model registry is: a field added to `RowProperties` or `ColumnProperties` with no
accessor fails the build naming the field, because otherwise the record would carry it, the codecs
would read and write it, and the public handle would simply never mention it.

Those sixteen pairs are spelled out rather than generated from a facet list, and the decision was
taken deliberately once the list-driven form was costed. The facet lists elsewhere in `core/`
(`CELL_STYLE_FACETS`, `WORKSHEET_MODEL_FACETS`) drive *loops* over state a caller never names;
generating accessors is a different thing, and its price is paid twice. `gen-docs.ts` reads class
members, so each property would still need a declared member carrying its doc comment, leaving only
the four-line body to save. And the members would have to be installed at runtime rather than
declared, which trades a surface the compiler checks for one it merely believes. That is the
`AssertNever` proof above spending its own guarantee. A slice that is not one costs more than the
lines it removes, and so does an abstraction.

The plumbing *underneath* those accessors was costed separately and also declined, which is worth
stating because it looks like the cheaper half of the same idea. Four members are byte-identical
between the two handles modulo which coordinate they name: the private read and write helpers, and
the `values` getter and setter. Lifting them into a shared module needs the pair of stores the
handle reads through, `peek` (which never fabricates) and `ensure` (which materialises on first
write), as an object the handle holds. That object and its two closures are then allocated per handle, and a
handle is constructed on every `getRow`/`getColumn` and once per step of `rows()`/`columns()`, so
iterating twenty thousand rows and reading one property measured about 17% slower. Formatting is
created on write and never on read precisely so a handle costs nothing; paying three allocations to
build one contradicts that.

The allocation-free shape is a base class, and it fails the other constraint: `gen-docs.ts` reads
class members, so `values` moving to a base would drop out of the reference unless `Row` and
`Column` redeclare it, at which point nothing is shared. Written out, the free-function version came
to thirty-two lines inserted against thirty-two removed, plus a sixty-line module. The read and
write helpers became pass-throughs that exist only to spell the store's name, and the `values`
setter became a closure forwarding to the call it replaced. The repetition is four small methods;
the abstraction was four small functions behind an interface, and it did not read better.

The same test was applied to `DataValidationOverlay` and `ConditionalFormattingOverlay` and reached
the same answer for a different reason. About fifteen lines are byte-identical between them: an
`add` that clones, an `entries` getter, a `shift` that maps every entry through `shiftSqref` and
drops what comes back `undefined`, and a `clear`. A shared base needs a type parameter, an accessor
for whichever field holds the sqref, and an injected clone function, and even then the two are not
symmetric: `DataValidationOverlay` also maintains a decoded-rectangle index that `at()` answers
point lookups from and that `shift` must re-derive, and the conditional-formatting overlay has
neither. A base class would carry one subclass's collection discipline while the other's sat outside
it, which is the slice-that-is-not-one test failing again.

If that is ever taken up, the honest shape is not a base class but a free function both `shift`
methods call, `shiftSqrefEntries(entries, refOf, withRef, axis, start, count, delta)`. That is the
part which is genuinely one rule, and it is the part where a drift would silently re-aim a
validation or a highlight at cells nobody chose.

**Costed and declined once, on the grounds that the rule was better placed elsewhere.** The rule at
risk is "a `sqref` whose every area the splice deleted takes its entry with it", and that rule lives
in `shiftSqref`, not in either caller: `src/core/merge.test.ts` now pins it there, along with the
byte-clean guarantee that an unmoved area comes back as its own text. What the free function would
extract from the callers is three lines, and it would need a `refOf`/`withRef` pair only because the
two spell the field differently (`sqref` against `ref`), while the validation overlay would still
keep its loop to re-derive the rectangle index. More indirection than it removes; taken up only if a
third `sqref`-bound overlay appears.

The xlsx reader and writer, the two largest pieces here, are each a cluster rather than a
monolith, split along the OOXML package's own divisions so a change touches one part:

- **read** (`src/io/xlsx/`): `read-styles.ts` (`styles.xml`), `read-worksheet.ts` (one sheet),
  with `read.ts` keeping `readXlsx` and the workbook-level wiring. `rich-runs.ts` owns the
  `<r>`/`<rPr>`/`<t>` element machine the worksheet and shared-strings parsers share, taking the
  container name (`is` or `si`) as a constructor argument, since that is the only thing that differs
  between a rich string Excel pooled and the same string written inline; `cell-accumulator.ts` owns
  the per-cell gathering state machine the buffered and streaming readers both drive (ADR-0004).
  Both are *machines*, not bags of state calls, and for one reason: a grammar two readers spell out
  separately is kept in step by convention rather than by mechanism, and the drift it admits reads
  the same content two ways depending on which encoding the producer happened to choose. Every other
  parser that gathers an element's text across open/text/close events does it through `TextCapture`
  in `src/xml/xml-read.ts`, which exists because nine hand-rolled copies each had to remember that a
  self-closing `<x/>` fires no close and a latch nothing closes is a latch that eats the next
  element's text. `capturedText` is its pull shape, for a parser whose whole job is reading a handful
  of text elements out of a part; a parser that interleaves capture with per-element state of its own
  stays bespoke.
- **write** (`src/io/xlsx/`): `package-plan.ts` (the part-graph plan layer), `workbook-xml.ts`
  and `worksheet-xml.ts` (the serialisers), `relationships.ts` (the SpreadsheetML relationship-type
  vocabulary), with `write.ts` keeping `writeXlsx` and the `buildPackageParts` orchestrator.

Neither cluster owns the container it rides in. `src/io/opc/` holds what is true of *any* OOXML
package: `inflate.ts` (the bounded inflater), `sniff-format.ts` (the magic-byte probe and the
typed rejection), `read-opc.ts` (resolving relationships and walking a part closure), `rels.ts`
(emitting a `.rels` part, escaping its own attributes rather than asking each caller to),
`part-paths.ts`, and the package namespaces. Beneath even that,
`src/xml/` holds escaping, emission and the SAX reader.

### The write boundary is where a value stops being a value and becomes bytes

A `PageSetup`, a `Color`, a `ConditionalFormattingRule` and a dozen shapes like them are plain data
the model stores verbatim, with no accessor to validate through. That is deliberate: making them
classes to catch a mistake would change the public surface of every one to guard a moment that has
not happened yet, and a `Color` that never reaches a file never had a problem. So the check belongs
at the single point where the value becomes bytes, and `src/xml/xml.ts` states all three forms of it:

- **A number is refused if the format cannot spell it.** `assertWritableNumber`, and the
  `numberText` / `numAttr` pair that call it. Every numeric attribute in OOXML is `xsd:double`,
  `xsd:unsignedInt` or a bounded flavour, and no lexical space has a form for a NaN or an infinity,
  so writing one produces a package Excel reports as damaged. The refusal is also exported on its own
  because a value can be unwritable and still be *read* on the way to the bytes: the collapsed-row
  scan compares outline levels to walk a group, and against `-Infinity` every comparison holds and
  the walk never ends.
- **A token from a closed enumeration is checked, not escaped.** `checkedToken`, against the
  `is*` guard in `src/core/` that owns that enumeration. Escaping a bogus token yields a well-formed
  document Excel still rejects, which buries the mistake in the file instead of raising it at the
  call.
- **A free string is escaped.** `escapeAttr` / `escapeText` / `textAttr`, plus
  `escapeSpreadsheetText` for the measured set of elements where Excel decodes `_xHHHH_`. A
  character XML cannot carry at all is refused, because there is no faithful representation and a
  sheet named `Sheet_x0001_A` is a different name rather than a rendering of the one asked for.

The read side is the same grammar with one asymmetry, and stating it is what makes the pair
trustworthy. `enumToken`, `numFinite`, `numInteger` and the `bool*` family in `xml-read.ts` *drop*
what they cannot read, where the writer throws. A file the library did not write is allowed to be
wrong, and losing one attribute beats losing the sheet; a value an author supplied is a mistake at
the call. Both halves lean on one guard per enumeration, so what the reader accepts is always
something the writer can write back. Where that leaves a model fragment unwritable anyway, the
reader drops the fragment: a `<cfRule>` whose type is not in `ST_CfType` is dropped whole rather than
half-read, because `type` is the attribute every other field is read relative to.

A **reference** is a foreign scalar like any other, and has its own tolerant pair in
`core/address.ts`: `tryDecodeCellRef` and `tryDecodeRange`, returning `undefined` where
`decodeCellRef` / `decodeRange` throw. Read-side code uses those two and nothing else, so a
malformed `r`, `ref` or `sqref` costs the element that carried it and never the sheet: the cell is
skipped, and the validation, hyperlink, table, filter or note is dropped. They answer "can this
name something that exists", not merely "does this parse", because `A0` and a row past the last
parse cleanly and then throw at the grid, which is the same abort one step later. Bounds are part
of the question, so they are part of the answer. An axis a range omits stays unbounded rather than
unreadable: `A:A` is a legitimate reference, and a caller needing a bounded rectangle checks the
corners itself.

That leaves each feature to decide what "drop" means for it, and both readings are in the tree.
A table is dropped *whole* when its `ref` is unreadable, for the `<cfRule>` reason: the anchor is
the coordinate every other field is read relative to. A `sqref` is dropped *per area*, because its
areas are independent and one unreadable area says nothing about its neighbours. Where the
authoring path shares the code, the guard stays on the authoring side and the reader filters before
reaching it, so `addDataValidation` still refuses a range naming no cells while a file carrying one
simply loses that validation.

## How a failure is reported

Every error the library raises deliberately descends from `XlsxError` (`src/errors.ts`), so one
`catch` clause answers "was that us?" without naming a class. Two levels of branch sit under it,
chosen so neither is redundant with the other: `code` is the *kind* of failure, `name` (and
`instanceof`) is exactly which one. Several classes share a code on purpose. A code in 1:1
correspondence with the classes would carry nothing the class did not already carry.

| code | what the caller does about it | classes |
| --- | --- | --- |
| `unsupported-format` | try a different reader, or reject the input | `UnsupportedFormatError` (its `format` field says *which* unsupported input) |
| `malformed-input` | reject the file, which is either broken or hostile | `PackageReadError`, `XmlParseError`, `XlsxParseError`, `XlsbParseError`, `VbaParseError`, `CustomUiParseError` |
| `authoring` | fix the calling code, which described a document that cannot exist | `AuthoringError`, `VbaAuthorError` |
| `internal` | report it, because an invariant of ours did not hold | `InternalError` |

Scalar argument validation stays outside the taxonomy: an index out of range, an unparseable
reference, a value of the wrong type are native `RangeError` / `SyntaxError` / `TypeError`, because
that is what those types are for. The line is composite against scalar: `getColumn(0)` is a
`RangeError`, while a table that names a column twice is an `AuthoringError`. A layer that wraps a
lower-level failure passes it as `cause` rather than flattening it into the message, except where
the lower layer's text is itself the hazard (a zip library's message can name an absolute path, so
`sniff-format.ts` replaces rather than wraps it).

There is deliberately no "not implemented yet" code. Every candidate turned out to be an
unreachable exhaustiveness guard, and the one real feature gap, that a binary `.xlsb` cannot be
row-streamed, is already reported through `UnsupportedFormatError`'s `format` branch.

## Two serialisations, one model

`.xlsb` is not a second library bolted on; it is a second codec over the same `Workbook`. The two
formats share an OPC/ZIP container, a relationship graph, and a style model, and differ only in how
the office-document parts are spelled: XML in `.xlsx`, BIFF12 record streams in `.xlsb`. The code
follows that split exactly, and the directory layout states it. The bounded inflater, magic-byte
probe and OPC/relationship resolution live in `src/io/opc/`, and the resolved-format table
(`XfStyle`, its built-in number formats, and `applyXfToCell`) in `src/io/style/`. Both sit *above*
the codecs rather than inside either. Only the part parsers live apart in `src/io/xlsb/`:
`record-stream.ts` (the record framing), `primitives.ts` (RkNumber, length-prefixed strings,
colours), `formula.ts` and `ptg-functions.ts` (the Ptg token stream a binary formula is stored as,
decoded to the text `<f>` would have carried), then a per-part parser mirroring its XML counterpart.

`readXlsx` detects which serialisation a package holds and dispatches, so a caller never branches on
format. The property that keeps the two honest is asserted, not assumed: the corpus reads one workbook
Excel saved in *both* forms and requires the two models to be identical. Anything the binary states
that XML omits must therefore be dropped on the binary side: a bottom vertical alignment, a locked
cell, a row restating the sheet's default height, a hatch fill's automatic-colour sentinels. That is
where most of the subtlety in that reader lives.

Namespace URIs and ext-URI GUIDs are registered once, split by which layer owns them: the
package-level URIs (`.rels`, content types, the relationship vocabulary) in
`src/io/opc/namespaces.ts`, the SpreadsheetML and extension ones in `src/io/xlsx/namespaces.ts`.
Sheet-local relationship ids are handed out by a single monotonic `SheetRelIds` allocator. Id
prefixes were once re-derived by hand-summing every prior part's count, which silently collides two
parts onto one id when a prefix drifts. Ids are now unique by construction and never recomputed by
arithmetic.

The public API is seven curated entry barrels under [`src/entries/`](../src/entries/), one
per subpath the package publishes (`/core`, `/xlsx`, `/xlsb`, `/csv`, `/vba`, `/customui`,
`/errors`), plus [`src/index.ts`](../src/index.ts), which unions them so the bare package name
still carries everything. Each symbol is listed in exactly one entry, so the root barrel is a
union of `export *` lines rather than a second list to keep in step.

That disjointness is load-bearing rather than tidy: `export *` does not report an ambiguous
re-export, it silently drops the name, so a symbol exported from two entries would vanish from
the root specifier with no diagnostic anywhere. `scripts/check-entries.ts` (a gate in
`verify --full`) fails the build on a duplicate, on an entry `package.json` does not publish, and
on a published subpath whose module is gone. It is also why the *whole* failure taxonomy is
exported from `/errors` and nowhere else: a container-level failure belongs to no single codec,
since `readXlsx` and `readXlsb` both raise `UnsupportedFormatError`, so putting the classes with the
codecs would have forced exactly the duplication the union cannot survive. That entry costs 12 KB,
so classifying a failure never loads a parser.

The barrels are curated, not exhaustive: modelled core-feature types (autofilter, page setup,
sheet views, defined names, image options) are public, and internal helper functions stay off
them. `src/vba/index.ts` and `src/customui/index.ts` are *internal* barrels that the model and
the codecs import; the public `/vba` and `/customui` faces are deliberately narrower. The two
halves of the streaming API are symmetric in reach but asymmetric in what needed naming.
Everything the streaming *writer* exposes is public: its workbook/worksheet/row handles are
classes, its options are interfaces, and nothing structural is left un-named. The streaming
*reader*'s per-row/cell output stays inferred-structural rather than a named commitment.
Streaming is not its own subpath: measured, it reaches every module `/xlsx` does plus three, and
an entry that costs what the codec costs is an alias rather than a boundary.

`"sideEffects": false` is declared, and it is true. No module in `src/` mutates anything at
import time, so a bundler may drop an unused one whole. `scripts/size-budget.ts` measures each
subpath's static-import closure against its own budget, which is what notices a boundary being
crossed: a codec acquiring a value-import of something it previously needed only as a type moves
one of those numbers and leaves the package total untouched. `scripts/smoke-dist.ts` resolves
every subpath through the package name, which is self-reference and the only thing in the repo
that exercises the `exports` map at all, then asserts `/core` reaches no serialisation and
`/errors` reaches nothing but itself.

Those numbers say one thing about the model that is worth stating rather than leaving to be
rediscovered: roughly half of `/core`'s weight is the VBA codec and the ribbon parser, because
`core/workbook.ts` imports `parseVbaProject`, `addVbaReference` and `removeVbaModule` at *value*
level. The layering rule permits that deliberately, since a workbook models a VBA project and its
ribbon, so those types are part of the document, and the measurement is what shows the price of
letting the operations live there too. It is left alone knowingly. Removing it means a split the
model does not import at value level, which breaks `Workbook`'s VBA methods, and it would pay off
only for a consumer importing `/core` with no bundler at all; with a bundler, `sideEffects: false`
already prunes what such a consumer never calls. Revisit it when that consumer turns out to exist
rather than on the strength of the number alone. Until then `./core`'s budget is what notices the
weight growing further.

The [API reference](api/README.md) is generated straight from the root barrel, so it cannot
describe a shape the compiler wouldn't accept.

## The VBA subsystem

Macro-enabled workbooks (`.xlsm`/`.xltm`) carry their VBA as a single opaque part,
`vbaProject.bin`, an OLE2 / Compound File ([MS-CFB]) container holding Office
run-length-compressed ([MS-OVBA]) module source and p-code. OOXML treats it as a binary
blob referenced by a workbook relationship; the ZIP package is otherwise identical to a plain
`.xlsx`. `src/vba/` is the self-contained, dependency-free subsystem that reads that blob and
applies pure-TS structural edits natively, with no `fflate` and no runtime dependency, just bytes.
Authoring or editing module *source* is not here: it needs genuinely compiled p-code only a real
Excel can emit, so it lives in the offline `tools/vba-compiler` (see below, and ADR 0019).

**Preservation is the safety floor, and every VBA feature is additive over it.** Read captures
`vbaProject.bin` (and its relationship/content-type closure, including a sibling
`vbaProjectSignature`) as a `PreservedWorkbookReference`, and the writer re-emits those bytes
verbatim, so a load/edit/save of an `.xlsm` keeps its macros with no VBA-specific code on the
common path. Everything below layers onto that guarantee; none of it can desync the two
representations, because a *read-and-not-re-authored* project is still emitted from the
preserved bytes alone.

The subsystem is built as encode/decode pairs over two formats plus a project layer:

- **CFB container.** `cfb.ts` reads and `cfb-writer.ts` writes. The reader walks the
  header, FAT, directory and mini-FAT, and reconstructs the whole storage/stream tree so the
  edit path can re-emit it with one stream swapped. The writer emits a v3 container whose
  storages are the name-ordered balanced red-black tree a *navigating* host (Excel) needs, not
  just the linear scan our own reader would accept.
- **MS-OVBA compression.** `ms-ovba.ts` (`decompressContainer` / `compressContainer`) is the
  chunked copy-token/literal-run codec Office uses for module source and the `dir` stream. It
  is *not* deflate. The compressor's contract is that its output re-expands byte-for-byte.
- **Project layer.** `project.ts` decodes the `VBA/dir` and module streams into a typed view;
  `project-editor.ts` splices structural edits (remove module, add reference) into an existing
  project; `vba-encoding.ts` holds the shared `dir`-record TLV encoders and VBA name validation.
  `codepage.ts` handles the project code page (MBCS, not latin1) in both directions, and
  `errors.ts` holds the two failure types. There is deliberately no from-source synthesizer;
  see below.

**The reader is hostile-input-facing (CLAUDE.md §3).** Every CFB sector index, chain, and
stream size is bounds-checked and cycle-guarded; every MS-OVBA back-reference is validated and
total output is bomb-capped. A malformed project fails closed with `VbaParseError`, never a
crash, hang, or unbounded allocation, and each guard is pinned by a crafted-malformed fixture.
The authoring/encode side is *our own* bytes, so it fails closed with `VbaAuthorError` on a
contract violation (over-long or duplicate stream name, unrepresentable character) rather than
emitting a silently broken container.

The public API layers by fidelity and intent, each slice fail-closed:

- **Read.** `Workbook.vbaProject: VbaProject | undefined` parses the preserved bytes *lazily*
  and memoises. Modules expose `name`, `streamName`, `kind` (the full procedural / document /
  class / designer classification), and decompressed `source`. A read never perturbs what the
  writer emits (ADR 0016).
- **Attach, replace, strip.** `Workbook.vbaProjectBytes` is a get/set accessor pair over the
  raw blob. A set is validated by `parseVbaProject` *before* any state change, so a bad blob is
  rejected whole, and replacing or removing drops the old bytes' now-stale signature. This is also
  how an authored project is installed: the offline `tools/vba-compiler` produces a compiled
  `vbaProject.bin`, and a consumer attaches it here, in pure TS with no Office at runtime.
- **Structural edits, pure TS.** `removeVbaModule` and `addVbaReference` (with their `Workbook`
  and package-level `editXlsxVba*` wrappers) splice the original `.bin`. They edit only the `dir`
  stream (and, for a removal, `PROJECT`/`PROJECTwm`) and leave every module stream and
  `_VBA_PROJECT` byte-for-byte. They are safe *precisely because* they never touch a module's
  compiled p-code: the `dir` stream, authoritative for the module/reference list, carries the
  change (ADR 0018/0019).
- **Author or edit module source, offline and not in the shipped library.** This is done by
  `tools/vba-compiler`, which drives a real headless Excel through the VBIDE object model to emit
  genuinely compiled, source-matched p-code (a `vbaProject.bin`, or a whole edited `.xlsm` for
  document code-behind). It is a Windows+Excel build tool, never in CI, whose output seeds
  committed fixtures (ADR 0019).

**Why source authoring cannot be pure TS.** Excel does not recompile VBA from source on open. A
module runs the compiled p-code (PerformanceCache) it ships, and the source is only recompiled
when a human opens the VBE. A `vbaProject.bin` synthesized from source alone, with no p-code or
mismatched p-code, either throws "Invalid data format" or *silently runs stale code*, so "opens
clean" is not proof of correctness. Only a real Excel can produce runnable p-code, hence the
offline compiler. Authoring artifacts are verified with `execute-verdict.ps1`, which opens with
macros enabled and *runs* a known authored macro, not merely by an open verdict (ADR 0019).
*Executing* macros in-process stays out of scope forever. That needs a live host, and this is a
document tool, not a VBA interpreter (ADR 0013).

## Tech decisions

The stack is deliberately small and each choice is recorded as an ADR under
[`docs/decisions/`](decisions/):

- **Runtime and no-build dev path.** ADR-0001.
- **Toolchain.** oxlint for rules and oxfmt for layout, with type-aware rules from tsgolint,
  ADR-0036; `node --test` over Vitest, hand-rolled type-level
  tests. ADR-0029. The `test/` and `scripts/` trees are TypeScript held to the same strict bar
  as `src/`, gated by `typecheck:test`, ADR-0011. All of those gates run as one concurrent,
  content-cached command (`node scripts/verify.ts`), ADR-0022.
- **Zip and XML write path.** `fflate`, plus a hand-written SAX reader with bounded allocation
  on every parser path. ADR-0003.
- **Docs generated from the types.** ADR-0006.
- **Spec reference.** Vendored OOXML schemas plus Microsoft Learn MCP. ADR-0007.
- **VBA subsystem.** Read view, ADR-0016; pure-TS structural edits, ADR-0018/0019; source
  authoring moved to the offline `tools/vba-compiler` after the "recompile cookie" premise was
  retracted, ADR-0019 (ADRs 0017/0018's from-source mechanism is retracted). Excel as a test
  oracle, ADR-0013.

## Working agreements

- **Preserve provenance as knowledge, not as a link.** Capture the real-world scenario a
  bug taught us. That survives; an upstream issue number does not. Durable artifacts
  (corpus cases, spec notes, commit messages) never cite upstream numbers; the commit
  that lands a change is its account of record.
- **Security- and correctness-first.** Every parser path is hostile-input-facing: no
  unbounded allocation, no zip-bomb naïveté. Entities are decoded but never expanded;
  inflation is bounded by a running output counter, not any declared size.
- **On a round-tripping surface, ask whether input is safe to *write back*.** Bounding what a
  parser will hold is only half of it. What the reader accepts, the writer re-emits, so a value
  a foreign part carries can leave *our* output invalid, and one invalid attribute is enough for
  Excel to offer to repair the feature away. A wire bound therefore lives in the parser, the
  authoring verb, *and* the serialiser, so the serialiser cannot emit an illegal value however the
  model was populated (`MENTION_OFFSET_MAX` is the worked example; see
  `knowledge/specs/threaded-comments-and-the-legacy-fallback.md`).
- **Two parts that are two halves of one representation derive from one variable.** Some features
  are only coherent as a pair. A threaded comment's conversation part and the legacy fallback
  `<comment>` that binds a cell to it are each *invisible in Excel* without the other, even though
  either alone validates clean. The writer computes such a pair from a single source (`write.ts`
  derives both from one `threads` value) so neither can be emitted without the other by
  construction, rather than by a rule someone has to remember.
- **A part family graduates from preserved to modeled in one change, never both at once.**
  Byte preservation (`core/preserved.ts`) is the sole emission authority for what it covers, so a
  *read view* over preserved bytes is safely additive (`Workbook.vbaProject`, `Workbook.customUI`,
  `loadedPivotTables`). But the moment a serialiser emits a part from the model, that rel type
  must leave `isPreservedSheetRelType`/`isPreservedWorkbookRelType` in the same commit or the
  package carries it twice.
- **Narrow foreign tokens; never trust them into the model.** An enumerated attribute
  read from a file is admitted only through a type guard that recognises the known union
  members (the pattern is `isCustomFilterOperator` in `core/autofilter.ts`); an
  unrecognised token is *dropped*, leaving the facet unset, rather than cast in with `as`.
  The reader's posture is *skip, never guess*: a malformed token yields absence, not a
  bogus value the rest of the code will trust. See ADR-0004 for the read path this serves.
  Nothing gates this the way `check-layering.ts` gates the import graph, and the closest
  candidate is declined on the record: `no-unsafe-type-assertion` finds 50 in `src/`, 45 of
  them `arr[i] as T` escapes that `noUncheckedIndexedAccess` and the ban on `x!` force, and it
  cannot tell one of those from a token. This rule is enforced by review, so a guard is what a
  reviewer looks for; the enumerations each carry one (`isCustomFilterOperator`, `isVisibility`,
  `isThemeColorSlot`, `isDeclarablePivotSourceKind`), keyed off a `Record` over the union where
  the compiler can then refuse an omitted member.
- **No half-migrations on main.** Each change is fully green, meaning typed, linted, tested and
  corpus-passing, and leaves the tree better than it found it.
