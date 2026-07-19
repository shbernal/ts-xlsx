# Architecture

How `ts-xlsx` is built and why. `CLAUDE.md` is the constitution (the principles every
change answers to); this document describes the library as it stands and the working
agreements that keep it coherent. Point-in-time decisions live under
[`docs/decisions/`](decisions/) as ADRs.

## Origin

`ts-xlsx` is a hard fork of [ExcelJS](https://github.com/exceljs/exceljs), which had
gone effectively unmaintained while still serving tens of millions of downloads a month.
The two assets trapped in that project were of opposite kinds: **knowledge** — thousands
of hours of hard-won understanding of how real-world `.xlsx` files behave, scattered
across hundreds of issues and PRs — and **code**, a weakly-typed, callback-flavored tree
with a rotting dependency graph. The strategy was to separate them: harvest the knowledge
into a durable, implementation-blind form, then rebuild the code from scratch against it.

That harvest is complete. Every credible bug, reproduction, and edge case became a
[regression corpus](../test/corpus/) case; the legacy tree is deleted; the runtime
dependency is now [`fflate`](https://github.com/101arrowz/fflate) alone. What remains is
a modern, strict-TypeScript library whose correctness is pinned by the corpus it carried
across.

## The corpus is the spine

The [regression corpus](../test/corpus/) is the product's backbone. Each case encodes
"correct behavior" as **implementation-blind** assertions that run against any
implementation through a thin adapter — so a behavior, once captured, can never silently
regress. This is why the corpus outlived the rewrite: it was written against the
*behavior*, not the code, so it validated the new implementation the same way it
indicted the old one.

The rule that follows: **when in doubt, add a case.** A bug without a corpus case is a
bug that will return. A missing feature is best reported as a corpus case so it is fixed
once and never regresses.

## Module layout

The source tree under [`src/`](../src/) is strict-TypeScript, ESM-only, and build-free on
the dev/test path (Node runs the `.ts` sources directly via type-stripping; `tsc` is the
type *checker* and, for publishing, the emitter). The domain decomposition, in dependency
order:

| Area | Role |
| --- | --- |
| core model | `Workbook` / `Worksheet` / `Row` / `Cell`, addresses, styles — the in-memory document |
| xlsx read/write | OOXML parse and serialize; the hardest, highest-value surface |
| streaming | bounded-memory row streaming for reads |
| csv | a thin, optional entry point, never coupled to the xlsx core |

The public surface is a single curated barrel, [`src/index.ts`](../src/index.ts). The
[API reference](api/README.md) is generated straight from it, so it cannot describe a
shape the compiler wouldn't accept.

## Tech decisions

The stack is deliberately small and each choice is recorded as an ADR under
[`docs/decisions/`](decisions/):

- **Runtime & no-build dev path** — ADR-0001.
- **Toolchain** (Biome for lint/format; `node --test` over Vitest; hand-rolled
  type-level tests) — ADR-0002.
- **Zip & XML write path** (`fflate`; a hand-written SAX reader with bounded allocation
  on every parser path) — ADR-0003.
- **Docs generated from the types** — ADR-0006.

## Working agreements

- **Preserve provenance as knowledge, not as a link.** Capture the real-world scenario a
  bug taught us — that survives; an upstream issue number does not. Durable artifacts
  (corpus cases, spec notes, commit messages) never cite upstream numbers; the commit
  that lands a change is its account of record.
- **Security- and correctness-first.** Every parser path is hostile-input-facing: no
  unbounded allocation, no zip-bomb naïveté. Entities are decoded but never expanded;
  inflation is bounded by a running output counter, not any declared size.
- **No half-migrations on main.** Each change is fully green — typed, linted, tested,
  corpus-passing — and leaves the tree better than it found it.
