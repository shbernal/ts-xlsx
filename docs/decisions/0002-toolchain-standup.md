# ADR 0002 — Toolchain standup: Biome for lint/format, `node --test` kept, tsc for type tests

**Status:** Accepted (2026-07-19) · Phase 4 · resolves the toolchain deferrals from [ADR 0001](./0001-rewrite-runtime-and-toolchain.md)

## Context

ADR 0001 deferred three tooling decisions until `src/` was large enough to justify
their config: a linter/formatter (Biome), a test runner (the `STRATEGY.md` default
was Vitest), and type-level tests. The rewrite is now ~76 modules with 647 unit
tests and a 671-behavior corpus, all green. The publishable build shipped in the
prior slice ([ADR 0001 addendum](./0001-rewrite-runtime-and-toolchain.md)). It is
time to stand the toolchain up — and to decide each piece on its merits now that
there is real code to measure against, rather than inherit the pre-rewrite defaults.

## Decision

### Biome is the lint + format toolchain

One binary, zero transitive dependencies, covers TS and the `.mjs` corpus harness.
`biome.json` mirrors the style the tree was hand-authored in (2-space, single quote,
semicolons, `bracketSpacing: false`, trailing commas, 100-col), so adoption was a
near-no-op format pass rather than a restyle. Linter runs Biome's `recommended`
preset. Scripts: `lint` (`biome check`), `lint:fix`, `format`. `lint` is now the
first gate in `npm test`.

Two deliberate config choices:
- **`noNonNullAssertion` is disabled for test files only** (an `overrides` block
  matching `**/*.test.ts` and `test/**/*.mjs`). In a test a non-null assertion is
  the honest idiom — the fixture is known, so `getCell('A1')!.value` beats
  defensively narrowing a value the test itself just created. Production code keeps
  the rule (CLAUDE.md §2 prefers narrowing over "trust me" escapes); the one `src`
  site was fixed by using `String.prototype.charAt` (typed `string`, no assertion).
- **The unsafe autofixes were applied, then verified against the corpus.** Biome
  classes `useOptionalChain` and `useTemplate` as unsafe (they *can* change
  semantics). We applied them and leaned on the full gate (typecheck + 647 unit
  tests + 671 corpus behaviors) to prove behavior was preserved. One unsafe fix was
  reverted by hand: `noSparseArray` rewrote the intentional sparse literals
  `['x', , 'z']` in two `addRow` hole-skipping tests to `['x', undefined, 'z']` — a
  real semantic change (`1 in arr` is `false` for a hole, `true` for `undefined`).
  Those two sites now carry a `biome-ignore` with the reason.

> **Gotcha, load-bearing:** Biome 2.5.4 **silently drops the entire `overrides`
> array when `biome.json` contains `//` comments.** The config parses, the rest of
> it applies, but overrides vanish — so the test-file rule relaxation evaporates and
> ~50 warnings reappear. `biome.json` is therefore kept as **comment-free JSON**;
> rationale lives here instead. Do not add comments to it. (CI would catch the
> regression as a wall of warnings, but the cause is deeply non-obvious.)

### `node --test` is kept; Vitest is rejected

The `STRATEGY.md` default was Vitest. We reject it. The 647 tests already run green
under Node's built-in runner on `.ts` sources with zero build step, and everything
Vitest would buy us we already have without its dependency tree:
- **Coverage** → `node --test --experimental-test-coverage` (script `test:coverage`).
- **Type-level tests** → tsc (see below), not `expectTypeOf`.
- **Watch** → the inner loop is already build-free; `node --test --watch` exists.

Vitest is a large transitive surface (Vite, esbuild, rollup, chokidar, …) — exactly
the kind of dependency weight the fork exists to shed (CLAUDE.md §2, §4). It earns
its place only if we need a capability Node's runner lacks; today we don't. Tests
stay in `node:assert/strict` style, so if that day comes the migration is mechanical.

### Type-level tests are hand-rolled, checked by tsc

`src/type-tests/` holds `Expect<Equal<…>>` assertions over the public barrel — no
`expectTypeOf`/`tsd` dependency. The standard bivariance-safe `Equal` and an
`Expect<T extends true>` are ~10 lines (`expect.ts`); the assertions
(`public-api.type-test.ts`) lock contracts the runtime tests cannot see: that
`writeXlsx`/`readXlsx` are synchronous (not `Promise`-returning), that a
`CellAddress`'s `col`/`row` are `number | undefined`, that `CellValue` never admits
`undefined`, and that every public symbol still exports. They are enforced by the
existing `typecheck` gate (`tsc` over all of `src/`) and excluded from the published
build (`tsconfig.build.json`). A drifted contract fails typecheck — the type-level
analogue of a red test, verified here by a negative control.

## Consequences

- **Positive:** lint + format is one zero-dep binary; the whole authored tree
  (src + scripts + corpus) is consistently formatted and lint-clean; no runtime-runner
  dependency added; type contracts are now regression-guarded; `npm test` gained a
  `lint` gate up front and `test:coverage` is available on demand.
- **Negative / deferred:** Biome's `recommended` is the whole ruleset — no
  project-specific rules curated yet; coverage is available but **not** yet a CI
  gate with a threshold (deliberate — a threshold is a follow-up once we decide the
  floor); the Biome comment gotcha is a sharp edge documented but not fixable by us.
- **Revisit when:** Node's test runner blocks a capability we need (reconsider
  Vitest — the migration is pre-planned); coverage should become an enforced floor;
  or the ruleset needs curation beyond `recommended`.
