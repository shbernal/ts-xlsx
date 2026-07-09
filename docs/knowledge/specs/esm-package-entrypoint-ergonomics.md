# The primary import must construct correctly on every toolchain

Cluster: packaging / DX

## Scenario

A developer installs the library, writes the most natural import for their toolchain, and
constructs the top-level workbook type. Depending on bundler and module system — a
webpack-backed React+TypeScript app, a vite/remix build, a nuxt/nitro server, plain Node ESM,
or a browser bundle — the constructor comes back `undefined` and instantiation throws
"X.Workbook is not a constructor", or the imported namespace is an empty object. The only way
forward is trial-and-error import shapes: reaching through a `.default` property, destructuring
the named class off a default import, swapping namespace vs default import syntax, or toggling
bundler caching. The same class name that works in one project is unconstructable in another
purely because of module-interop packaging — nothing to do with spreadsheet data.

## Desired behavior

The primary public API must be importable identically and predictably across every mainstream
toolchain, with the ergonomic import being the correct one. A single documented form —
`import { Workbook } from '<pkg>'` — must construct correctly under Node ESM, TypeScript
(NodeNext/bundler resolution), and the common bundlers (webpack, vite, rollup, esbuild, remix,
nuxt/nitro). No `.default` reach-through, no empty-namespace failure mode, and no divergence
between what the types say is importable and what exists at runtime.

## Root cause

The ecosystem failures stem from shipping a CommonJS-only entry whose CJS/ESM interop places
the real exports under a synthetic `default`. Bundlers and Node's ESM-over-CJS interop then
disagree about whether the top-level named export exists, producing `undefined` constructors or
empty namespaces. The converged workarounds (`pkg.default.Workbook`, `const { Workbook } = pkg`,
default-vs-namespace swaps) are symptoms of that packaging, not fixes the library should require.

## How the fork precludes this

Ship real ESM as the source of truth with named exports. If CJS is supported at all, provide a
correct dual export map (`exports` with `import`/`require`/`types` conditions); otherwise drop
CJS entirely per the fork's ESM-only stance and eliminate the interop seam. Ensure the emitted
`.d.ts` matches the runtime shape exactly so `import type` and value imports agree. Lock it down
with type-level export-surface tests plus a small matrix of import-and-construct smoke checks
across resolvers/bundlers, so a regression cannot ship. The plain named import must be the easy
path; no interop workaround should ever be necessary. This directly serves the constitution's
"ESM, modern idioms only, correct use is the easy path".

## Open questions

- Does the fork support CJS consumers at all, or is it strictly ESM (which makes the dual export
  map unnecessary and removes the interop seam entirely)?
- Which resolver/bundler matrix is the committed support surface for the import smoke tests?
- Single default export, named exports only, or both — reflected consistently in the export map
  and the type declarations?

Related notes: `public-types-node-stream-portability`, `browser-safe-io-boundary`,
`no-global-polyfill-in-browser-bundle` — the same "the package must behave predictably wherever
it is consumed" principle applied to types, IO, and browser globals.
