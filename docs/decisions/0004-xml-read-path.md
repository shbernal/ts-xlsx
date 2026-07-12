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
  processing instructions, `xml:space`) — covered by unit tests. The inflate bound uses
  the zip's *declared* uncompressed size; a header-lying bomb (declares small, inflates
  large) needs the streaming inflate with a running byte counter, which lands with the
  streaming reader slice. This is documented, not silently assumed.
- **Revisit when:** a real-world file exercises an XML construct the lean parser does
  not cover (record it as a corpus fixture first), or the streaming reader needs the
  hard, running-counter inflate bound.
