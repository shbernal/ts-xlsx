# Strategy

> The phased plan that operationalizes `CLAUDE.md`. This document is **living** —
> agents update it as phases complete, decisions are made, and reality diverges
> from the plan. If you change course, change this file in the same breath.

**Working name:** `ts-xlsx` (npm: `@shbernal/ts-xlsx`) — a **provisional** handle chosen
2026-07-09 to make the JS/TS ecosystem and file format legible. It is not the final
brand; the definitive rebrand remains a human decision (see Phase 4). Keep it a single,
easily-changed field (`package.json#name`) — do not sprinkle it through the code.

---

## The core insight

The two assets locked inside upstream ExcelJS are of **opposite** kinds:

1. **Knowledge** — thousands of hours of hard-won understanding of how real-world
   `.xlsx` files behave, encoded as bug reports, reproductions, PRs, and edge-case
   fixes across ~650 issues and ~140 PRs. **This is priceless and must be preserved.**
2. **Code** — a 20k-line Babel-transpiled, callback-flavored, weakly-typed JS
   codebase with a rotting dependency tree. **This is debt and will be discarded.**

The strategic error would be to conflate them: to either (a) rewrite from scratch and
*lose the knowledge*, or (b) keep the codebase to *keep the knowledge*. We refuse both.

> **The plan: harvest the knowledge into a durable, implementation-independent form
> (specs + a regression corpus + a triaged backlog), stabilize the current code
> just enough to validate that corpus, and only then hard-fork into a modern
> TypeScript rewrite that must pass the corpus we carried across.**

The regression corpus is the bridge. It is written to be independent of *any*
implementation, so it survives the rewrite and proves the new code is at least as
correct as the old — plus everything the old one got wrong.

---

## Backlog snapshot (captured at fork time)

| Metric | Value |
|---|---|
| Upstream last `master` commit | 2024-01-12 |
| Upstream last release | v4.4.0 (2023-10-19) |
| Open issues | ~654 (76 `bug`, 17 `help wanted`, 8 `enhancement`, 7 `proposal`, 4 `Typescript`, …) |
| Open PRs | ~139 (72 `fix`, plus feat/type/style/stream) |
| PR hotspots (by title keyword) | table (11), type (8), style (7), stream (5), pivot/image/conditional (4 each), csv (3) |
| Current runtime deps to replace | `archiver`, `jszip`, `unzipper` (zip); `saxes`+`sax` (xml); `fast-csv`; `dayjs`; `uuid`; `tmp`; `readable-stream` |
| Current source | ~171 JS files, ~20.4k lines under `lib/` |
| Existing hand-written types | `index.d.ts` (~48 KB) — a *harvest asset*, not authoritative |

Regenerate this snapshot with the harvest tooling (Phase 0) rather than trusting it
to stay current.

---

## Architecture map (what exists → what it becomes)

The current `lib/` layout is a reasonable domain decomposition and informs — but does
not dictate — the target module boundaries.

| Current area | Lines | Role | Target treatment |
|---|---|---|---|
| `lib/doc` | ~4.4k | In-memory model (Workbook, Worksheet, Row, Cell, styles) | Redesign as a strict typed core domain model; the heart of the rewrite |
| `lib/xlsx` | ~11.4k | OOXML read/write (the `xform` transformers) | Reimplement against a new XML layer; the hardest, highest-value area |
| `lib/stream` | ~2.1k | Streaming read/write | First-class async iterators / Web Streams, not `readable-stream` |
| `lib/csv` | ~0.4k | CSV in/out | Thin, modern, optional entry point |
| `lib/utils` | ~2.0k | Shims & helpers | Mostly deleted; replaced by platform + tiny focused modules |

---

## Target tech stack (decisions, not suggestions)

These are the default decisions. Revisit only with a recorded ADR under
`docs/decisions/` if evidence demands it.

- **Language:** TypeScript, `strict` + `noUncheckedIndexedAccess` +
  `exactOptionalPropertyTypes`. ESM-only source and output. `.xlsx` core has no
  hard Node dependency — runs in Node, Deno, Bun, browsers, and edge runtimes.
- **Compression (zip):** **`fflate`** — replaces `archiver` + `jszip` + `unzipper`
  in one small, fast, zero-dependency, isomorphic library. Streaming zip in/out.
  This single swap kills most of the transitive-dependency CVE surface that drove
  the fork.
- **XML:** a small, fast, streaming-capable parser. Evaluate `fast-xml-parser` vs a
  purpose-built SAX layer (we control the schemas, so a lean hand-rolled reader may
  win on both size and speed). Decide with a benchmark, record as an ADR.
- **CSV:** a minimal modern parser (or a tiny dependency) behind an optional entry
  point — never coupled to the xlsx core.
- **Dates:** native `Temporal`/`Date` math; drop `dayjs`. Excel's 1900/1904 serial
  date system handled explicitly and tested to death (see harvested date bugs).
- **IDs:** `crypto.randomUUID()`; drop `uuid`.
- **Temp files:** eliminate `tmp`; stream in memory / to caller-provided sinks.
- **Build:** `tsup`/`unbuild`-class bundler emitting ESM + types. No Babel, no Grunt,
  no Browserify. Delete `.babelrc`, `gruntfile.js`, `benchmark.js`-as-is.
- **Test:** **Vitest** (fast, TS-native, coverage built in). Retire Mocha/Chai/Jasmine/Grunt.
- **Type tests:** `expectTypeOf` (Vitest) / `tsd` for the public surface.
- **Lint/format:** **Biome** (single fast tool) unless a required rule forces
  ESLint-flat + Prettier. No `prettier-eslint` Rube-Goldberg chains.
- **CI:** typecheck + lint + unit + regression corpus + `npm audit`/OSV +
  bundle-size budget, all required. Matrix across Node LTS + Bun + browser (via
  Playwright) for the isomorphic core.

---

## Phases

Each phase has an explicit **exit criterion**. Do not start a phase's dependents
until its exit criterion is met. Within a phase, agents parallelize freely.

### Phase 0 — Foundation & harvest tooling  *(enables everything)*
Set up the machinery before doing the work at scale.
- Stand up the TS toolchain, CI skeleton, and this repo's dev ergonomics.
- Build **harvest tooling**: scripts that pull all upstream open issues + PRs with
  bodies, labels, reactions, linked files, and attached sample `.xlsx` into a local,
  queryable dataset under `docs/knowledge/backlog/` (JSON + fetched fixtures).
- Define the **regression corpus format** under `test/corpus/` — each case is
  `{ description, provenance (issue/PR #), input fixture?, expected behavior }`,
  runnable against *any* implementation via a thin adapter.
- **Exit:** one end-to-end corpus case (harvested from a real upstream issue) runs
  red/green against the *current* code through the adapter.
  → ✅ **Met (2026-07-09):** issue #140 harvested via `scripts/harvest/fetch-issue.mjs`;
  corpus case `0140-address-decoding` runs red/green through the `current` adapter
  (`npm run corpus`). Remaining Phase 0 bullet: the additive CI skeleton. Toolchain
  rip-out stays deferred (see `PROGRESS.md`).

### Phase 1 — Harvest the backlog  *(preserve the knowledge)*
Convert the backlog from GitHub threads into durable project assets. This is the
phase that "leaves no opportunity behind us."
- **Triage & cluster** all open issues and PRs into themes (tables, styles,
  streaming, pivot tables, images, conditional formatting, dates, formulas, CSV,
  types, security/deps). The snapshot above seeds the clusters.
- For every credible bug/repro: distill a **corpus case** (failing test + fixture),
  tagged with its provenance. Reactions/severity set priority, not inclusion —
  we capture broadly, implement selectively.
- For every open PR: extract the *intent and the test*, not the diff. A PR's lasting
  value is the reproduction and the domain insight; the patch itself is written
  against code we are deleting. Record: problem, root cause, correct behavior,
  corpus case. Note which PRs are trivial dep bumps we simply supersede.
- For proposals/enhancements: capture as **spec notes** under `docs/knowledge/specs/`
  (desired behavior, prior art, open questions) feeding Phase 2 design.
- Treat the harvest as a **one-time drain**, not a ledger (see
  `docs/knowledge/BACKLOG.md`). `harvest:list` freezes the universe into
  `manifest.json`; `harvest:all` fills the queue (`backlog/issues/*.json`); agents
  then drain it thread by thread. **Distill the knowledge, delete the record, and let
  the commit message be the durable account** of what was preserved (or why an item
  was not carried). We do **not** maintain a per-item `{captured|superseded|…}` table,
  and durable artifacts never cite upstream numbers — they go meaningless post-fork.
  Nothing is silently dropped: the frozen `manifest.json` is the denominator, `git
  log` the per-item account, an empty queue the completion proof (`CLAUDE.md` §"no
  silent caps"). Follow it with `harvest:status`.
- **Exit:** the queue (`backlog/issues/`) is empty — every item drained; every carried
  item left a corpus case and/or spec note behind it. The corpus is large, runs
  against current code (mostly red where bugs are real), and is implementation-blind.

### Phase 2 — Stabilize-to-validate  *(prove the corpus, then let go)*
A *deliberately time-boxed* pass on the **current** code — its only purpose is to
validate the corpus and produce a correctness baseline, not to polish code we will
delete.
- Fix the high-value, low-risk bugs where doing so on current code sharply clarifies
  the correct behavior (turning corpus cases green teaches the rewrite the target).
- Land the security/dependency fixes that are cheap and de-risking, **if** they don't
  cost more than the rewrite would anyway.
- Do **not** invest in refactoring, typing, or restyling the legacy code. It is
  scaffolding.
- **Exit:** the corpus expresses a clear, agreed "correct behavior" for each captured
  case, with current-code pass/fail recorded as the baseline the rewrite must beat.

### Phase 3 — The rebuild  *(discard the debt)*
Greenfield TypeScript implementation, corpus-driven, module by module.
- Build order follows value and dependency: **core model → XML layer → xlsx
  read/write → streaming → csv**. Land each module fully green (typed, tested,
  corpus-passing) before the next depends on it — no half-migrations on main.
- The rewrite is "done" for an area when it passes **every captured corpus case**
  for that area, including the ones the legacy code failed.
- Design the public API fresh for humans + agents: discoverable, precisely typed,
  correct-by-construction. The old API and `index.d.ts` are *reference inputs* to
  this design, never constraints on it.
- **Exit:** feature-parity-or-better on the corpus for all in-scope areas; legacy
  `lib/` deleted; dependency tree small and audit-clean.

### Phase 4 — Independence & identity  *(the clean break)*

> **Update (2026-07-09): hosting independence pulled forward.** Rather than remain a
> GitHub fork until the end, we did the *infrastructure* clean break up front: a fresh,
> non-fork repo (`shbernal/ts-xlsx`) was created by mirror-pushing full history into it,
> inherited upstream branches were pruned to `master`, and package/license/README
> identity was repointed. `upstream` (exceljs/exceljs) is kept only as a **read-only**
> remote (push disabled) to feed the Phase 1 harvest, and gets removed once harvest is
> done. What remains deferred here — deliberately, per the drift lesson — is the *final
> rebrand name* and all *code modernization*, which stay last so the legacy tree keeps
> its shape as a mergeable base while we drain the backlog.

- **Stop tracking upstream.** Remove the `upstream` remote from the workflow. *(Now: it is
  already push-disabled and read-only; drop entirely at harvest end.)*
- Rebrand: new package name, docs, and identity (**human decision** — this is one of
  the few things escalated per `CLAUDE.md` §3). *(`ts-xlsx` is a provisional working
  name only; the definitive brand is still open.)*
- First-class docs generated from the types; migration notes framed as "this is a
  different, better library," not a compatibility shim.
- **Exit:** a `0.x` release of the independent library, published under its own name,
  with the corpus as its living correctness guarantee.

---

## Working agreements for agents on this plan

- **The corpus is the product's spine.** When in doubt, add a case. A bug without a
  corpus case is a bug that will return.
- **Preserve provenance as durable knowledge, not as a link.** Capture the *real-world
  scenario* a thread taught us — that survives the fork; the issue/PR number does not.
  Durable artifacts (corpus cases, spec notes, commit messages) never cite upstream
  numbers. The commit that drains an item is its account of record.
- **Capture broadly, implement selectively.** Phase 1 hoards knowledge cheaply;
  Phases 2–3 spend effort where value is highest. Reactions and real-world frequency
  guide priority.
- **Update this file.** Phase transitions, ADRs, and scope calls land here or under
  `docs/decisions/`. A stale plan is worse than no plan.

## Immediate next actions (Phase 0 kickoff)

> **Sequencing guardrail (2026-07-09):** the legacy tree's *shape* is frozen until the
> backlog is drained. High-drift moves — whole-tree reformat, `.js`→`.ts` rename,
> module-layout changes, dependency swaps — are what make the ~139 open PRs unmergeable,
> so they come **last**, not first. Phase 0 does only the *additive* work below (harvest
> tooling + corpus) that does not touch the legacy shape. See `PROGRESS.md` for live
> status.

1. ✅ Harvest toolchain built: `harvest:list` (universe → `manifest.json`),
   `harvest:all` (resumable queue fill), `harvest:status` (drain progress), plus the
   single-thread atom. Bodies, labels, reactions, attachments, and sample `.xlsx` land
   under `docs/knowledge/backlog/`.
2. ✅ Corpus format + adapter defined; one real harvested case runs end-to-end,
   red/green against the *current* code.
3. **Fill then drain:** run `harvest:all` once to fill the queue, then drain it per the
   `harvest-triage` skill — distill each thread into a corpus case and/or spec note,
   delete the record, commit. Follow with `harvest:status --clusters`.
4. **Deferred to the end (not now):** replace the Babel/Grunt/Mocha toolchain with the
   TS/Vitest/Biome/tsup skeleton. This is the highest-drift action in the whole plan;
   it runs only once the backlog is captured and the mergeable PRs are banked.
