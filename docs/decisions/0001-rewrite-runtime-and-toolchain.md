# ADR 0001 — Rewrite runtime & toolchain: run `.ts` directly, defer the bundler

**Status:** Accepted (2026-07-11) · Phase 3 kickoff

## Context

`STRATEGY.md` decided the rewrite stack up front: strict TypeScript, ESM-only,
Vitest, Biome, and a `tsup`/`unbuild`-class bundler emitting ESM + `.d.ts`. Those
were sound defaults written before the rewrite began. On starting Phase 3 two facts
changed the cheapest correct path for the *first* modules:

1. The dev environment runs **Node 24**, which executes `.ts` files directly via
   type-stripping — no transpile step, no loader, and an `.mjs` can `import` a
   local `.ts` module.
2. The corpus is the product's spine and already runs via plain `node`. The rewrite
   only needs to be *reachable* from a corpus adapter and *type-checked*; it does not
   need to be *bundled* to be proven correct.

Standing up Vitest + Biome + a bundler now is real dependency and config surface for
no correctness gain on a single pure module, and it is the highest-drift work in the
plan (`PROGRESS.md` guardrail). Deferring it keeps the first slices lean and
dependency-clean (`CLAUDE.md` §2).

## Decision

- **Runtime/test path uses Node's native `.ts` execution — no build step.** The
  `rewrite.mjs` corpus adapter imports `src/**/*.ts` directly; unit tests run under
  the built-in `node --test` runner on `.ts` files. Local TS imports use explicit
  `.ts` extensions.
- **`tsc` is the type-safety *gate*, not a build tool.** `npm run typecheck`
  (`tsc --noEmit -p tsconfig.build.json`) enforces the full strict flag set
  (`strict`, `noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`,
  `noImplicitOverride`, `verbatimModuleSyntax`, …). TypeScript is pinned to 5.x.
- **The new tree is ESM, scoped by `src/package.json` (`"type": "module"`).** The
  legacy CommonJS root (`lib/`, `excel.js`) is left byte-for-byte intact — no
  root `"type"` flip that would reinterpret legacy `.js` as ESM.
- **Vitest, Biome, and the ESM/`.d.ts` bundler are deferred** to a dedicated
  toolchain-standup slice, to run once enough of `src/` exists to justify the
  packaging/lint config and after the legacy Grunt/Babel/Mocha rip-out is scheduled.

## Consequences

- **Positive:** zero-build inner loop; no new runtime/test deps beyond `typescript`;
  legacy tree untouched so the freeze guardrail holds; strict typing still enforced.
- **Negative / deferred:** no published bundle or emitted `.d.ts` yet (fine — nothing
  is published in Phase 3); two test runners transiently (`node --test` for `src/`,
  Mocha for legacy `spec/`) until the swap; `node --test` assertions are written in a
  Vitest-portable style (`node:assert/strict`) to make that migration mechanical.
- **Revisit when:** a target runtime without `.ts` execution must run the code, a
  publishable artifact is needed (Phase 4), or the `src/` surface is large enough that
  Biome's lint/format and Vitest's watch/coverage pay for their config.
