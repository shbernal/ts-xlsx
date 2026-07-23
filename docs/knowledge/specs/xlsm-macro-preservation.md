# Preserve VBA on round-trip of macro-enabled workbooks (.xlsm)

Cluster: formats

## Scenario

A user maintains a macro-enabled spreadsheet (.xlsm) with VBA automation — for example a report
template whose embedded macro exports to PDF. They want to open it, edit worksheet data
programmatically, and save it back with the macros intact so it still runs when reopened in Excel.
Today, reading and rewriting such a workbook silently strips the VBA project, yielding a plain .xlsx
that has lost its macros. (The related, explicitly out-of-scope wish is to *invoke* macros from
Node; that needs a live Excel/automation host and is not a document-tool concern.)

> Spec note, not a corpus case: capturing this needs a real .xlsm fixture and a feature that does not
> yet exist, so there is no assertable behavior to lock today. The durable value is the preservation
> contract and its packaging details.

## Desired behavior

- On reading an .xlsm, the library **retains the VBA project part** (`vbaProject.bin`) and its
  associated relationships and content-type declarations for **byte-faithful round-trip** — the
  preserved bytes stay the sole thing the writer emits.
- On writing, if the source carried a VBA project (or the caller explicitly requests macro-enabled
  output), the package is emitted with the correct macro-enabled content type and the
  `vbaProject.bin` part **re-embedded byte-for-byte** — so a round-trip that only edits cell values
  yields a file Excel still treats as macro-enabled with functioning macros.
- Callers can **read** the macro source through a lazy, read-only typed view
  (`Workbook.vbaProject`), derived from the preserved bytes without a write-back path — see
  ADR 0016. Reading never perturbs what the writer emits.
- A signed VBA project's `vbaProjectSignature` part is preserved alongside — it is a sibling part
  reached from `xl/_rels/vbaProject.bin.rels`, so the closure walk carries it through, and its
  distinct `.bin` content type is re-declared per-part (not collapsed into the `vbaProject` default,
  which a single extension `<Default>` would otherwise do — locked by `preserved-parts.test.ts`).
  Editing the *project itself* legitimately invalidates the signature; editing unrelated cells does
  not (see open questions).
- **Out of scope:** executing/running macros, and *authoring/editing* VBA (read-only for now — see
  ADR 0016). The library is a document tool, not a VBA interpreter.

## Prior art

OOXML defines a distinct macro-enabled content type for the workbook part and packages the VBA as a
binary `vbaProject.bin` referenced by a workbook relationship; the ZIP container is otherwise
identical to .xlsx. Many OOXML libraries treat the VBA project as an opaque blob they copy through
without parsing. Macro-enabled templates (.xltm) are the template analog.

## Open questions

- API surface: automatic preservation whenever a VBA part is detected on read, or gated behind an
  explicit flag when writing?
- Output-format choice (.xlsx vs .xlsm): infer from source, from filename extension, or an explicit
  option?
- ~~Expose the VBA project bytes to callers, or only pass them through opaquely?~~ **Decided
  (ADR 0016):** exposed as a lazy, read-only typed view (`Workbook.vbaProject`), derived from the
  preserved bytes with no write-back path. Authoring (source → bytes) stays deferred pending a
  forcing consumer.
- How strictly to handle/warn about the signature part when the workbook is mutated. (The part now
  survives round-trip correctly typed; the *policy* on drop-vs-warn when the project is edited still
  waits for a consumer. An `isSigned` accessor is sourceable from the preserved closure once needed.)
- Parse macro/toolbar-referenced `customUI`, or leave it opaque? (Audit whether it survives
  round-trip today — likely yes via the generic preserved-parts net, but unproven.)

Related: `roundtrip-preserves-unmodeled-package-parts`; ADR 0016 (read view + authoring deferred).
