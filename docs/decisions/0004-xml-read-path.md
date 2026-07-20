# ADR 0004 — The XML read path is a lean, hand-written SAX pull parser

**Status:** Accepted (2026-07-12) · Phase 3 reader slice

## Context

ADR 0003 stood up the writer on `fflate` and emitted XML by direct string assembly,
deliberately **deferring the XML *parser* choice to the reader slice**. That slice is
now here: almost every remaining corpus capability (`roundtripWorkbook`,
`readFixtureReport`, styles/defined-name/merge read-back — ~48 cases) asserts through a
package that is *read back*, so the model cannot light them up without a reader.

The deferred decision was `fast-xml-parser` (a DOM-building dependency) versus a lean
hand-written SAX. Two forces from the constitution (`CLAUDE.md`) decide it:

1. **A spreadsheet reader parses untrusted input.** "No unbounded allocation, no
   zip-bomb naïveté, no eval-shaped surprises." A DOM parser materialises the *entire*
   part as an object tree before the reader sees a single cell — an allocation
   multiplier an attacker controls, exactly the shape we must not ship.
2. **The dependency tree stays small, modern, and clean** — a founding reason for the
   fork. A parser is not a dependency we want to own transitively when the subset of
   XML that OOXML uses is small and regular.

There is also a build-order force: the eventual **streaming** reader (large files,
`.eachRow`) fundamentally needs a pull/SAX model, not a DOM. Building SAX now is the
same primitive that path will extend, not throwaway work.

## Decision

- **The reader is a lean, hand-written SAX pull parser** (`src/io/xlsx/xml-read.ts`),
  no XML-library dependency. It emits open/text/close events in a single O(n) pass with
  no recursion; the OOXML reader (`src/io/xlsx/read.ts`) consumes them and builds only
  the model, so peak memory tracks real content, not document structure.
- **Entities are decoded, never expanded.** Only the five predefined entities
  (`&amp; &lt; &gt; &quot; &apos;`) and numeric character references (bounded to valid
  Unicode) are recognised. DTDs and `<!ENTITY>` definitions are ignored outright, so
  *billion-laughs* entity expansion and XXE external-entity resolution are structurally
  impossible — not mitigated, absent.
- **Inflate is bounded** at the reader's entry (`readXlsx`): a cap on total declared
  uncompressed size rejects the naïve zip bomb before the parser runs.
- **`fflate` serves both directions.** Reading uses `unzipSync` (with the size filter);
  the write path already used `zipSync`. One zip library, as ADR 0003 set out.

## Consequences

- **Positive:** zero new dependencies for the read path; the hostile-input guards live
  in code we own and test; the SAX primitive is what the streaming reader will extend;
  escaping (write) and entity-decoding (read) are each one small audited surface.
- **Negative / deferred:** we own XML edge cases (attribute quoting, CDATA, comments,
  processing instructions, `xml:space`) — covered by unit tests.
- **Revisit when:** a real-world file exercises an XML construct the lean parser does
  not cover (record it as a corpus fixture first).

## Update (2026-07-13) — the inflate bound no longer trusts declared sizes

The original slice bounded inflation by the zip's *declared* uncompressed size, and flagged
the gap: a header-lying bomb (declares small, inflates large) slips past, and — worse —
trusting the declared size to preallocate lets an attacker force a large allocation from a
few compressed bytes. Both are now closed by `src/io/xlsx/inflate.ts`: the package is fed to
fflate's streaming unzip in bounded slices, the decompressor grows its output from the bytes
it *actually* produces, and a running counter aborts the moment real output crosses the cap.
Declared sizes are consulted for nothing. `maxUncompressedBytes` now bounds produced output,
not header claims. This is the first slice of the streaming reader; the same streaming-inflate
primitive is what an eventual row-streaming (`.eachRow`) read path extends.

## Update (2026-07-13) — the pull parser and the streaming row reader

The SAX parser was push-only: `parseXml(source, handlers)` drove callbacks over the whole string
in one loop. A callback cannot `yield`, so a reader that must *emit* incrementally could not sit on
top of it. The scan loop is now extracted into a generator, `xmlEvents(source)`, that yields
`open`/`text`/`close` events; `parseXml` is a thin push adapter over it, so every existing call site
is byte-for-byte unchanged (the corpus proves it) while a pull consumer can now drive the parse.

On that primitive sits the first streaming *read* API: `readSheetRows(data, options)`
(`src/io/xlsx/read-rows.ts`), a generator that yields one worksheet's rows in order as plain
`{number, cells}` records, retaining only the row in hand rather than materialising the whole
`Workbook`. Value decoding is shared with the buffered reader through one module
(`src/io/xlsx/cell-value.ts`), so a cell streamed one row at a time decodes identically to the same
cell read as part of a full workbook — the divergence such a split would otherwise invite is closed
by construction. This slice still inflates the package whole (bounded as above) and reads shared
strings / styles as whole parts — both legitimately document-sized; a later slice can make the
inflate itself per-part lazy on the same pull primitive.

## Update (2026-07-20) — the reader consumes the pull parser through a small helper vocabulary

The hand-written SAX was right, but the *consumption* side had drifted into three habits worth
naming so they don't return: ~30 no-op `onText(){}`/`onClose(){}` stubs, ~20 inline
`=== '1' || === 'true'` boolean idioms in three incompatible forks, and every parser hand-coding
both a self-closing branch *and* an on-close branch for the same element. The read path now consumes
events through a settled set of helpers built on the same primitive, so a parser states only what it
means:

- **`SaxHandlers.onText`/`onClose` are optional** — a handler declares only the events it consumes.
- **`openElements(xml, ...localNames)`** (`xml-read.ts`) is a pull generator over `xmlEvents` for the
  "scan opens, read attributes" readers, collapsing an accumulator-threaded-through-a-closure into a
  plain `for..of` that reads `attrs` directly.
- **`closeEmptyElements(events, names)`** expands a self-closing `<x/>` whose local name is in `names`
  into open+close, so a formatted-but-empty element commits once in `onClose`. It is a **per-name
  opt-in** (`ReadonlySet<string>`), deliberately *not* a blanket "synthesize a close for every empty
  tag" mode: text-bearing elements (`<f/>`, `<v/>`, `<t/>`) commit *captured text* on close, so a
  synthesized close would misread them — a self-closing shared-formula clone `<f t="shared" si=".."/>`
  would set `hasFormula` and be mis-read as a shared-formula *master*. Only elements safe to run
  on-close-when-empty are listed.
- **Boolean attributes** go through `boolPresent` / `boolStrict` / `boolTristate`; **numeric operands**
  through `coerceNumericLiteral` (strict decimal regex `^-?\d+(?:\.\d+)?$` — a non-canonical spelling
  like `1E5` is kept *verbatim*, the round-trip-faithful and more conservative read of foreign input);
  a leading `=` through `stripFormulaEquals` (returns *unescaped* text — the caller escapes for its
  target).

**Enumerated attributes are narrowed, never trusted.** A union-typed attribute token is admitted only
through a guard that recognises the known members; an unrecognised one is dropped rather than cast in
with `as` (see the *narrow foreign tokens* working agreement in `docs/architecture.md`). This is a
behaviour change on *malformed* input only — valid tokens are unchanged, so the byte corpus stays
green; the drops are pinned by unit tests.

**Generator gotcha (load-bearing).** A sub-parser that drains a slice of a *shared* event generator
must pull with `.next()` (or a helper generator that `return`s), **never** `for..of` + `break` —
breaking out of a `for..of` calls the generator's `.return()`, which terminates the shared stream for
every later sub-parser. The style-table driver's `until(events, container)` is the canonical shape:
one `xmlEvents` pass, each section's sub-parser draining its own container's slice, no re-scan.

## Update (2026-07-20) — one cell-gathering state machine, two finalisers

The earlier update shared value *decoding* across the buffered and streaming readers; cell
*gathering* — the per-cell `<c>`/`<f>`/`<v>`/`<is>` state each reader accumulated — was still
re-implemented twice, free to drift. It is now one class, `CellAccumulator`
(`src/io/xlsx/cell-accumulator.ts`, modelled on the `rich-runs.ts` run accumulator): both readers
drive the same `beginCell`/`beginFormula`/`setFormula`/`setValue`/`appendText` surface, so a cell
can never be *gathered* differently depending on which reader saw it.

Finalisation, by contrast, deliberately diverges — and that split is the point. The buffered reader
calls `finalize`, which resolves shared-formula masters/clones and data-table cells and opens rich
`<r>` runs; the streaming reader calls `decode`, the plain-decode subset, which must **not** resolve
shared formulas (it surfaces the clone's own cached result) and must **not** open rich runs (a
streamed inline string flattens to text). `finalize`'s plain path is itself routed through `decode`,
so there is exactly one `RawCell`-build and one decode. Sharing gathering while forking finalisation
is what lets the two readers stay honest to their different contracts without duplicating the fragile
part. A malformed non-empty `<c r>` now throws in `beginCell` for *both* readers, closing another way
the two could have diverged.
