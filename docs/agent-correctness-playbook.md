# Agent correctness playbook

> One page, so an agent mid-task never has to reconstruct the decision tree. It maps
> **what you are doing** to **the check that proves it correct** and the exact command.
> The capabilities themselves are described in `docs/architecture.md` and the ADRs;
> this is the dispatch table on top of them.

The net is defense-in-depth. From cheapest/fastest to most authoritative:

| Layer | What it proves | Command | Needs |
| --- | --- | --- | --- |
| Types + unit | The code compiles under strict TS and units pass | `pnpm run typecheck && pnpm run test:src` | Node 24 |
| Lint | Style/format/floating-promise/console gates | `pnpm run lint` | Node 24 |
| **Corpus** | Well-formed XML, package structure, and no behavior regression — the **spine** | `pnpm run corpus` | Node 24 |
| **OOXML oracle** | Schema + semantic conformance vs Microsoft's own validator | `pnpm run validate:ooxml -- file.xlsx` | **.NET 10** |
| Spec grounding | Ground a decision in the authoritative format | Learn MCP + `schemas/` + `docs/knowledge/specs/` | — |

**Cost is not the only axis — authority is (ADR 0012).** These layers witness three
different things, and a lower one cannot stand in for a higher one:

- **Self-consistency** — a write→read round-trip is a fixed point of *our own* code. It
  catches *unilateral* writer/reader bugs; it is structurally blind to *correlated* ones
  (both halves wrong in compensating directions) and to anything spanning two package
  parts. Sufficient only for **intra-model** claims (a value survives, a style does not bleed).
- **Spec-conformance** — the `OpenXmlValidator` oracle and the `inspectPackage` structural
  facts. An *independent* implementation, so it breaks the round-trip correlation — but it
  enforces what ECMA-376 *states*, not what Excel *does*. Required for **single-part
  conformance**.
- **Excel behavior** — what Excel Desktop actually does. The only ground truth for cross-part
  invariants the spec omits (e.g. a table's header cells must exist and match its column
  names). On a Windows+Excel host, scriptable via the Excel-oracle harness for state-observable
  behavior (ADR 0013); recorded as `provenance: {source: 'excel-desktop-verification'}`.

For a **cross-part correspondence**, one Excel-Desktop verification *seeds* the invariant
and a corpus fact whose shape *is* the relationship *locks* it. The `inspectPackage`
vocabulary is partitioned by part and cannot phrase most cross-part seams yet — ADR 0012
lists the open ones.

`pnpm test` runs lint → typecheck → test:src → corpus. The **Stop hook** runs
typecheck + test:src + corpus automatically at each turn boundary *when `src/` is
dirty*, so you cannot end a turn green while regressing the spine. The OOXML oracle
is **not** in the hook (it needs .NET and is slower); invoke it yourself — see below.

## Situation → check

**You added or changed a writer path (anything that emits XML).**
Run `pnpm run corpus` — it parses the written package and asserts well-formedness,
part/relationship/content-type structure, and element ordering
(`test/corpus/adapters/ooxml-facts.ts`), plus every behavior regression. Then run the
schema/semantic oracle on a representative file — use the **`validate-ooxml` skill**,
which emits a workbook and runs `pnpm run validate:ooxml` for you. New behavior ships
with a corpus case in the same change (use the **`write-corpus-case` skill**).

**You added or changed a reader path (parsing foreign XML).**
Treat all input as hostile (ADR-0004): no unbounded allocation, no entity expansion,
inflation bounded by output counted, unrecognized tokens dropped — never cast with
`as`. Add a **fixture-backed corpus case** for any real-world file shape you learn
about (`test/corpus/fixtures/<case>/…`), then `pnpm run corpus`. A round-trip case
(write → read) is the strongest reader proof.

**You are fixing a bug.**
Test-first. Write an implementation-blind corpus case that reproduces it
(`write-corpus-case` skill), set its `baseline` to what the code does *today*, watch it
fail, then fix until `pnpm run corpus` is green. We never fix the same bug twice.

**You need a cross-part / Excel-quirk invariant seeded — the only ground truth is what Excel Desktop does.**
On a Windows host with Excel installed, don't do it by hand: run the Excel-oracle harness.
Write a probe (`tools/excel-oracle/probes/<invariant>.json`: a cell spec + the cells to
observe), then `node tools/excel-oracle/run.ts <probe.json> --out test/corpus/fixtures/excel-oracle/<invariant>.json`.
It opens the file headless over COM, reads formula/value per cell, re-saves to reveal the
geometry Excel considers canonical, and writes an auditable observation sidecar. This
**seeds** the invariant only — then **lock** it with a Tier-2 seam fact that runs in CI and
a case carrying `provenance: {source: 'excel-desktop-verification', ref: '<sidecar>'}`. The
harness is a probe, not a test: it needs Windows+Excel+`pwsh`, self-guards to a loud refusal
without them, and **never** runs in CI (`pnpm run corpus` must not depend on Excel). It answers
*state-observable* questions on *one Excel build* only — see [ADR 0013](./decisions/0013-excel-desktop-as-automatable-tier3-oracle.md)
for what is and isn't scriptable and the five standing pitfalls.

**You are unsure how an OOXML element / attribute / enum / child-ordering should look.**
Do not guess — the format is full of surprises. In order:
1. Read the vendored XSDs: `schemas/ooxml-transitional/`, start at `sml.xsd` and follow
   its imports (the authoritative element structure, types, enums, ordering). These are
   **read-only reference** — see the note below.
2. Query the **microsoft-learn MCP** (`microsoft_docs_search` / `microsoft_docs_fetch`)
   for Excel's *real-world deviations* from the standard — the prose the XSDs can't
   encode. This is enabled for the project (ADR-0007); if a run says the server isn't
   available, enable `microsoft-learn` for the project.
3. Check `docs/knowledge/specs/` for a note we already wrote on the same corner.

**You changed the build/emit path (`tsconfig.build.json`, import specifiers, a runtime reference type-stripping tolerates).**
The dev/test loop runs *stripped* `src/` `.ts`; consumers run *`tsc`-emitted* `dist/` JS — two artifacts that can diverge. `pnpm run build && pnpm run corpus:dist` runs the full behavioral corpus against the emitted JS (`CORPUS_TARGET=dist`), not just the `smoke:dist` round-trip. CI's `build` workflow does this on every PR; run it locally when you touch anything emit-shaped.

**You are about to finish a turn / open a PR.**
The Stop hook covers typecheck + test:src + corpus. For full CI parity also run
`pnpm run lint` and, if you have .NET 10, `pnpm run test:ooxml`. CI runs all three
workflows (`build`, `corpus`, `ooxml-validation`) regardless, so the oracle is always
enforced before merge even when you can't run it locally.

## The schema oracle, and what to do without .NET

`OpenXmlValidator` (the `validate:ooxml` / `test:ooxml` scripts, ADR-0002) is the
**single authoritative** schema/semantic check. It needs .NET 10. Exit codes:
`0` = every input clean, `1` = validation/package errors found, `2` = the tool could
not run. Known, tracked errors are baselined in
`test/ooxml-validation/allowed-errors.json`; a *new* error fails the gate and a *stale*
baseline fails it too, so keep that file honest when you fix or introduce a diagnostic.

If you don't have .NET 10 locally, do **not** reach for a second validator. The vendored
XSDs are deliberately **not** wired into an `xmllint`-style path (`schemas/README.md`,
ADR-0002): a naive XSD-only pass gives false alarms and false confidence — it can't do
the semantic checks or validate the OPC parts, and the Transitional schemas are subtly
permissive. Instead: rely on `pnpm run corpus` (well-formedness + structure) locally,
read the XSDs and the Learn MCP to reason about correctness, and let CI's
`ooxml-validation` workflow run the authoritative oracle on your PR.
