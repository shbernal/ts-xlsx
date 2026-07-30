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
| **OOXML oracle** | Schema + semantic conformance vs Microsoft's own validator | `node scripts/ooxml-validator.ts file.xlsx` | **.NET 10** |
| Spec grounding | Ground a decision in the authoritative format | Learn MCP + `schemas/` + `docs/knowledge/specs/` | — |

**`lint:fix` needs no confirming `lint` pass.** `biome check --write` applies what it can and
*still exits non-zero* if any diagnostic survives, so a green `lint:fix` already is the proof.
Re-running `lint` after it only re-checks a tree you have been told is clean.

**A warning fails the gate** — every Biome invocation here passes `--error-on-warnings`, because
Biome exits 0 on warnings and most of the `style` group (`noNonNullAssertion` among them) is a
warning. When one fires, fix the code, do not reach for the autofix: `?.` on an assertion that was
load-bearing turns a crash into silent wrong output. Non-null assertions are usually a signal that
an index is being carried where the object itself could be — see `src/vba/cfb-writer.ts`.

**Run one corpus case, not 265, while you iterate.** `node test/corpus/run.ts --case
<id-or-cluster-glob>` is well under a second against ~13 s for the whole corpus, and prints the
case in full. `--json` gives one machine-readable report object. The summary line reaches stdout
in *every* mode — never pipe a run through `grep` to find a case, and never run the corpus twice
to get both the detail and the tally.

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

**To run the whole net at once, use `node scripts/verify.ts`** — every gate above plus
`docs:check` and `constitution:check`, run concurrently, reported as one table with
per-gate timing (~14 s wall against ~27 s of serial work). Prefer it over assembling the
chain by hand, which is how `docs:check` and `constitution:check` get silently dropped.
`--quick` is the inner loop: types, unit tests, and lint scoped to your changed files, no
corpus (~5 s) — faster, but **not** a substitute for the full run. Invoke it with `node`,
not `pnpm run`, to skip ~1 s of package-manager wrapper. `pnpm test`, lefthook's `pre-push`
hook and CI's `corpus.yml` are all the same full run — there is no second list of gates to
keep in step, so adding one here is a one-line change that CI picks up. Why it is shaped this way — the pool width, the
cache key, the incremental-`tsc` traps — is [ADR 0022](./decisions/0022-verification-is-one-cached-parallel-entrypoint.md).

The **Stop hook** runs `verify --full --cached` at each turn boundary, so you cannot end a
turn green while regressing the spine. `--cached` exits immediately when the working tree
is byte-for-byte what it was the last time this gate set passed — a *hit means proven*, not
skipped, because the key is the HEAD commit plus the full diff and every untracked file. A
turn that changed nothing verifiable costs ~0.3 s; one that changed anything pays the real
~13 s. The OOXML oracle is **not** in the hook (it needs .NET and is slower); invoke it
yourself — see below.

**Write scratch to `.tmp/`** — probes, dumps, generated workbooks, anything regenerable
(`$SCRATCH` and `$TMPDIR` both point there; CLAUDE.md makes it the rule). It is git-ignored,
so probing leaves `git status` clean *and* costs nothing at the turn boundary: an untracked
file anywhere else is part of the cache key and buys you a full re-verify.

## Situation → check

**You added or changed a writer path (anything that emits XML).**
Run `pnpm run corpus` — it parses the written package and asserts well-formedness,
part/relationship/content-type structure, and element ordering
(`test/corpus/adapters/ooxml-facts.ts`), plus every behavior regression. Then run the
schema/semantic oracle on a representative file — use the **`validate-ooxml` skill**,
which emits a workbook and runs `pnpm run validate:ooxml` for you. New behavior ships
with a corpus case in the same change (use the **`write-corpus-case` skill**).

**A generated file's *content* is right but its *layout* opens wrong** — a frozen header
row unpainted until you click it, a missing outline bar, no sheet selected.
Suspect an omitted **view-initialisation** fact before you suspect the data or the styles.
Excel writes `<bookViews><workbookView/>`, `tabSelected="1"` on exactly one `<sheetView>`,
and `outlineLevelCol`/`outlineLevelRow` on `<sheetFormatPr>` into every file it saves;
consumers lay the pane geometry and the outline bars out against them, so omitting them
leaves that layout uninitialised. Such a package is still schema-valid and still opens
without a repair prompt — **neither the oracle nor `open-verdict.ps1` will flag it**. The
writer emits all three unconditionally now (`DEFAULT_WORKBOOK_VIEW`, `src/core/workbook.ts`).
That they were omitted is certain; that any *one* of them causes a given paint glitch is
inference from the diff — the single-variable A/B that would isolate it was never run, and
the original report only ever reproduced under a window geometry we could not recreate. So
if this class of symptom recurs with all three present, the cause is elsewhere: reopen the
investigation rather than assuming it regressed here.
The general move for this whole class: round-trip your output through Excel's own `SaveAs`
over COM and diff the two packages. What Excel adds unprompted is what a consumer expects
to find.

**You are cutting a release.**
Bump `version` in `package.json`, cut `CHANGELOG.md`'s `## [Unreleased]` into the new
version's section, commit, and push — then let CI go green *before* tagging, because the
tag is what the release names and a tag that fails its own gates is the one thing you
cannot quietly redo. Tag `vX.Y.Z`, push it, and publish a GitHub release on it: that
release event is what publishes to npm (ADR-0026), authenticated by OIDC with no
credential in the repository. Rehearse first if you want — dispatch `publish.yml` from the
tag with `dry_run` on — but note the rehearsal reaches `npm publish --dry-run` only for a
version the registry does not already serve. If the publish job fails the "tag and version
must be the same claim" step, fix `package.json` and re-tag; do not weaken the check. If it
fails at `npm publish` with a **404** on a package that plainly exists, that is npm refusing
the OIDC identity, not a missing package: read the trusted publisher on npmjs.com and check
its repository, workflow filename (`publish.yml`) and environment (`npm-publish`) against
the job. A publish that has never once succeeded is far more likely misconfigured there than
here — do not start editing the workflow.

**You added or changed a reader path (parsing foreign XML).**
Treat all input as hostile (ADR-0004): no unbounded allocation, no entity expansion,
inflation bounded by output counted, unrecognized tokens dropped — never cast with
`as`. Add a **fixture-backed corpus case** for any real-world file shape you learn
about (`test/corpus/fixtures/<case>/…`), then `pnpm run corpus`. A round-trip case
(write → read) is the strongest reader proof.

**You just wrote an assertion that guards an invariant — prove it can fail.**
A test that cannot fail is worse than no test: it reports a guarantee nobody is holding. Break the
thing on purpose, watch the assertion fire, restore. This has caught real theatre more than once —
an archive-length check meant to pin two writers to the same compression level passed happily at
level 9, because below roughly 200 rows every level compresses a fixture identically; a
compile-time exhaustiveness proof is only a proof once you have added a field and seen it named in
the error. Applies to every mechanism in this repo that exists to catch a future mistake: the facet
registries, the entry/layering gates, the size budgets. Cheap, and it is the difference between a
check and a comment.

**You are about to claim something is faster, or that it stops blocking the event loop.**
Measure, and distrust the first number — a bad measurement will talk you out of a correct change.
Three traps, all hit in one sitting while sizing `writeXlsxAsync`:
- **One process per case.** Running the baseline and the candidate in the same process loads the
  second with the first's GC pressure. On a ~42 MB payload that alone made the faster path look
  slower. Pass the case in on `argv` and run the script twice.
- **Never hand-roll a `setInterval` watcher for loop blocking.** It reports timer coalescing as
  blocking, and it reports `max = 0` when the loop is blocked so hard the callback never fires
  *once* — so the worst case reads as flawless. `perf_hooks.monitorEventLoopDelay` measures the
  actual thing.
- **Responsiveness and throughput are two claims.** Moving work to a worker can leave wall-clock
  untouched while cutting the longest stall from seconds to milliseconds. Say which one you
  measured; a change often buys one and not the other.

Probes go in `.tmp/`. Put the numbers in the commit or the ADR — the next agent should not have to
re-derive them to know whether the trade still holds.

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

**You added, removed or moved a public export.**
Symbols live in exactly one entry barrel under `src/entries/`, and `src/index.ts` is `export *`
over all seven — so adding a name in two places does not conflict, it makes the name *vanish* from
the root specifier with no error anywhere. `node scripts/check-entries.ts` (already in
`verify --full`) is what catches that, along with an entry `package.json` forgot to publish. Then
`pnpm run docs` — the reference is generated from the root barrel, so a symbol missing from the
diff is a symbol that fell out of the union. Error classes go in `src/entries/errors.ts` and
nowhere else (ADR-0023).

**You changed what a module imports, and it crossed a directory.**
`node scripts/check-layering.ts` proves the graph's direction still holds. If the import turned a
`import type` into a value import, also run `pnpm run build && pnpm run size`: the per-entry
closures are the only measurement that notices a codec joining an entry's graph — the package
total does not move when a boundary is crossed, only when something new is written.

**You are about to finish a turn / open a PR.**
The Stop hook covers every gate `verify --full` runs. For full CI parity add the one it
cannot: if you have .NET 10, `node test/ooxml-validation/run.ts`. CI runs all three
workflows (`build`, `corpus`, `ooxml-validation`) regardless, so the oracle is always
enforced before merge even when you can't run it locally.

## The schema oracle, and what to do without .NET

`OpenXmlValidator` (`scripts/ooxml-validator.ts`, also reachable as the `validate:ooxml` /
`test:ooxml` scripts, ADR-0002) is the **single authoritative** schema/semantic check. It
builds the .NET assembly on demand and then invokes it directly, so a warm call costs ~0.9 s
rather than the ~2–6 s `dotnet run` spends re-evaluating the project; validate several files
in one call. It needs .NET 10. Exit codes:
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
