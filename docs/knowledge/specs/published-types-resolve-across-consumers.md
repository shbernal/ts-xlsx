# Published types must resolve for TypeScript, ESM, and CommonJS consumers

Cluster: types

## Scenario

A developer installs the library and imports it from a TypeScript project, and separately from plain
ESM and CommonJS projects. They expect editor intellisense for the public API and expect `tsc` to
type-check consuming code against the shipped declarations. In some editor/toolchain combinations the
type information failed to resolve, so autocompletion and type checking silently degraded even though
the runtime import worked — the failure was environment-dependent (some editors, not others) and tied
to how the package manifest points at its entry and its declarations.

> Spec note, not a corpus case: this is a packaging/manifest guarantee verified by a build/CI
> consumer-smoke check and type-level tests, not by the runtime roundtrip adapter. (The upstream
> thread was a PR against a deleted Grunt/Babel/browserify build; the diff is discarded, the
> requirement kept.)

## Desired behavior

- The published package exposes its public API types so they resolve **correctly and identically**
  for three consumer shapes — TypeScript, ESM, CommonJS — regardless of `moduleResolution`
  (`node`, `node16`/`nodenext`, `bundler`). Editor intellisense and `tsc` both pick up the shipped
  declarations.
- An explicit `exports` map declares `types`, `import`, and `require` conditions so node16/nodenext
  resolution succeeds, with `types`/`typesVersions` fallbacks covering older resolvers.
- The `.d.ts` (or `.d.mts`/`.d.cts`) variants match each condition — no single hand-authored
  declaration that only resolves under one module system.

## Open questions

- Ship dual ESM+CJS, or ESM-only with a compatibility shim? Whichever is chosen, the declaration
  variants must line up with the runtime conditions.
- CI guard: install the built tarball into throwaway TS, ESM, and CJS consumer projects and assert
  both runtime import and type resolution succeed — turning "intellisense works" into a
  machine-checked invariant rather than a manual editor screenshot.
- Prefer type-level tests (`expectTypeOf`/tsd-style) over per-file compile-spec suites.

Related: `public-types-node-stream-portability`, `async-iterable-types-compile-cleanly`,
`write-buffer-return-type-contract`, `browser-safe-io-boundary`.
