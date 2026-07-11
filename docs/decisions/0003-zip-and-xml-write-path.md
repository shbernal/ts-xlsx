# ADR 0003 — Zip container is `fflate`; the write path emits XML directly

**Status:** Accepted (2026-07-12) · Phase 3 writer slice

## Context

The core in-memory model landed (`value`/`cell`/`worksheet`/`workbook`). The next
build-order step is the `.xlsx` writer, because almost every remaining corpus
capability (`inspectPackage`, `roundtripWorkbook`, …) round-trips through a written
package — the model cannot light those cases up on its own. Standing up the writer
forces two dependency decisions that `STRATEGY.md` left open:

1. **The zip container.** An `.xlsx` is an OPC zip. Legacy used `archiver` (write) and
   `unzipper`/`jszip` (read) — three overlapping, dated packages, part of why we
   forked.
2. **XML.** Legacy used `saxes`/`sax` for reading and hand-rolled string building for
   writing.

## Decision

- **Zip: `fflate`.** A single, tiny (~8 kB), zero-dependency, actively-maintained,
  ESM-first library that does both deflate *and* inflate. It replaces
  `archiver` + `unzipper` + `jszip` outright. The writer uses `zipSync`; the reader
  (later slice) will use `fflate`'s streaming inflate, which is where the
  hostile-input guards (bounded output, no zip-bomb naïveté) will live.
- **Writing XML needs no parser.** OOXML parts are emitted by direct string assembly
  through a tiny escaping/serialisation helper (`src/io/xlsx/xml.ts`). Correct
  escaping of text and attributes is the only hard requirement, and it is security-
  relevant (an unescaped `<`/`&`/`"` produces a malformed package a consumer rejects),
  so it is centralised and unit-tested rather than sprinkled inline.
- **The XML *parser* choice stays deferred** to the reader slice (`fast-xml-parser`
  vs a lean SAX). The writer does not need it, and deferring keeps this slice's
  dependency surface to exactly one new package (`fflate`).

## Consequences

- **Positive:** one modern dependency replaces three legacy ones; the same library
  serves both write and read; the write path has no XML-library dependency at all;
  escaping is one audited surface.
- **Negative / deferred:** the reader's parser decision is still open (tracked for the
  next slice); `fflate`'s `zipSync` is synchronous — acceptable for the buffered
  writer, and the streaming writer slice can move to its async/stream API later.
- **Revisit when:** the streaming writer needs incremental/async zipping, or the
  reader slice benching shows a different inflate is materially better.
