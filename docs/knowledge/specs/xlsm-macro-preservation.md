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
- Callers can **attach, replace, copy, or remove** a macro project at the bytes level through
  `Workbook.vbaProjectBytes` (a get/set accessor pair over the raw `vbaProject.bin`) — see ADR 0017.
  Setting bytes makes the written package macro-enabled and embeds them verbatim (validated fail-closed:
  a malformed blob is rejected, never half-applied); setting `undefined` demotes the workbook to a plain
  package. Replacing or removing drops a now-stale signature over the old bytes.
- Callers can **author a macro project from module source** through `Workbook.setVbaProject({modules})`
  (or the standalone `writeVbaProject`) — see ADR 0017 §2.3. It synthesizes a complete `vbaProject.bin`
  (CFB container + MS-OVBA-compressed `dir`/module streams) from a list of modules (name, kind, source);
  the modules carry no compiled p-code, so Excel recompiles them from source on open. Procedural and
  class modules are supported (document/designer are host-coupled and rejected fail-closed). **Verified
  against real Excel 365:** a synthesized workbook opens clean and, re-saved macro-enabled, preserves
  every module with its source. This makes the workbook an emission authority for authored macros — the
  reversal of ADR 0016's read-only-view core.
- Callers can **edit an existing module's source in place**, preserving the project's references,
  host-extender info, and every other module — see ADR 0018. The edit is a *splice* over the original
  `vbaProject.bin`: it replaces only the edited module's compressed source stream, zeroes its
  MODULEOFFSET, and resets `_VBA_PROJECT` to recompile from source, never touching the streams that carry
  references or other modules. Unlike from-scratch authoring, this **can edit document/designer modules**
  (e.g. `ThisWorkbook`, `Sheet1`) because it inherits their host linkage from the preserved streams rather
  than synthesizing it. Two surfaces: `Workbook.setVbaModuleSource(name, source)` (model level) and
  `editXlsxVbaModuleSource(xlsx, name, source)` / `editXlsxVbaModuleSources(xlsx, edits)` (package level,
  swapping only `xl/vbaProject.bin` and leaving every other part byte-for-byte). The package-level path is
  the highest-fidelity way to edit a real workbook — the model round-trip re-serializes the whole package
  and can perturb strict parts on rich files (a pre-existing, VBA-independent `writeXlsx` gap; see below).
  **Verified against real Excel** on a genuine 10-module workbook, including editing a document code-behind
  → opens clean and recompiles.
- A signed VBA project's `vbaProjectSignature` part is preserved alongside — it is a sibling part
  reached from `xl/_rels/vbaProject.bin.rels`, so the closure walk carries it through, and its
  distinct `.bin` content type is re-declared per-part (not collapsed into the `vbaProject` default,
  which a single extension `<Default>` would otherwise do — locked by `preserved-parts.test.ts`).
  Editing the *project itself* legitimately invalidates the signature; editing unrelated cells does
  not (see open questions). Both edit surfaces (ADR 0018) drop the now-stale signature on a source edit.
- **Out of scope:** executing/running macros (needs a live host — ADR 0013), and *adding or removing*
  modules/references (only editing an existing module's source is in scope — ADR 0018). Attaching or
  replacing a whole `.bin`, authoring a project from source, and editing an existing module's source are
  all in scope (ADRs 0017, 0018). The library is a document tool, not a VBA interpreter.

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
  preserved bytes with no write-back path.
- ~~Let callers author macros, or only preserve what a read produced?~~ **Decided (ADRs 0017, 0018):**
  the forcing-consumer gate is lifted and authoring is built out. Attach-blob (`Workbook.vbaProjectBytes`),
  from-scratch synthesis (`Workbook.setVbaProject`/`writeVbaProject`), and editing an existing module's
  source (`Workbook.setVbaModuleSource`, `editXlsxVbaModuleSource(s)`) all shipped, each fail-closed. Only
  *adding/removing* modules or references remains unbuilt (a future slice, no consumer yet).
- **Signature on a source edit is dropped** (ADR 0018): editing the project invalidates any signature over
  it, so both edit surfaces discard the stale `vbaProjectSignature` (part, relationship, content-type
  override) rather than leave it advertising a broken signature. Editing unrelated cells still preserves
  it (the part round-trips correctly typed). An `isSigned` accessor is still sourceable from the preserved
  closure once a consumer needs it.
- ~~Parse macro/toolbar-referenced `customUI`, or leave it opaque?~~ **Audited & fixed:** the ribbon
  parts (`customUI/customUI.xml`, `customUI14.xml`) hang off the *package root* `_rels/.rels`, not the
  workbook rels, so the workbook-closure net never reached them and the writer — which regenerates the
  root rels from the model — dropped them. Now captured as package-root preserved references and
  re-declared verbatim on write (locked by `preserved-parts.test.ts`). Parsing the ribbon XML into a
  model stays deferred (no consumer); it round-trips opaquely, like the rest of the preserved net.

Related: `roundtrip-preserves-unmodeled-package-parts`; ADR 0016 (read view + authoring deferred),
ADR 0017 (authoring in scope; attach-blob + from-scratch synthesis), ADR 0018 (editing an existing
module's source by splice).
