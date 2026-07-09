# The dependency tree must stay small, current, and audit-clean

Cluster: security/deps

## Scenario

A team runs a dependency audit in their build pipeline and it flags a moderate-severity advisory —
not in the spreadsheet library's own code, but in a **transitive** dependency several levels down (a
UUID generator pinned to an old major with a known buffer-bounds defect). Because the flagged package
is deep in the tree, consumers cannot upgrade it directly; they are forced into fragile package-manager
override/resolution tricks to make `audit` pass. They want the library itself to carry a dependency
tree so small and current that this never happens.

> Spec note, not a corpus case: there is no library behavior to assert implementation-blind — the
> issue is the shape of the dependency graph, checked in CI, not a runtime output. It is captured
> (rather than discarded as a trivial bump) because it encodes a durable design stance the fork
> already commits to. The specific advisory id and the old upstream pin are ephemeral and left out.

## Desired behavior

- The dependency tree is **small and modern by default** — every transitive package earns its place;
  a dependency audit is part of CI and is expected to stay green (a founding reason for the fork was
  upstream's rotting transitive dependencies).
- **Prefer platform built-ins over third-party packages** for commodity primitives. A UUID generator
  is the canonical example: use the runtime's built-in `crypto.randomUUID()` (Node ≥ 14.17,
  and the Web Crypto equivalent in browsers) instead of pulling a `uuid`-style dependency. Every
  eliminated dependency is one fewer advisory surface consumers can inherit.
- **No deep, un-upgradable pins.** The library must never wedge consumers into override/resolution
  hacks to satisfy their own audit; if a needed package has a vulnerable transitive, the library
  either upgrades, replaces, or vendors the minimal needed piece.
- Supply-chain hygiene is a **maintained invariant**, not a one-time cleanup — new dependencies are
  weighed against the built-in-first rule before adoption.

## Open questions

- Minimum supported runtime versions that let us rely on built-ins (`crypto.randomUUID`, Web Crypto,
  streams) without a polyfill dependency — and how that interacts with the browser build's
  no-global-polyfill stance.
- Where a browser target lacks a built-in a Node target has: conditional exports vs. a tiny inlined
  shim vs. requiring the host to provide it.
- CI policy: fail the build on a `moderate`+ advisory, or gate on `high`+ with a tracked allowlist?

Related: `no-unsafe-eval-csp-compatible`, `no-global-polyfill-in-browser-bundle`,
`esm-package-entrypoint-ergonomics`, `browser-safe-io-boundary`,
`sheet-protection-password-hash-compatibility`.
