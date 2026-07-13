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
