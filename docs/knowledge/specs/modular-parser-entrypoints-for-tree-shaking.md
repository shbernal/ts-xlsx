# Format loaders behind tree-shakeable entry points

Cluster: streaming

## Scenario

A browser or performance-sensitive application uses only a subset of the library's parsing
capabilities — it reads/writes only CSV, or only XLSX, or only manipulates an in-memory workbook
without touching any file format. Today all format loaders are pulled in transitively when the
workbook entry point is imported, so even a CSV-only consumer ships the entire XLSX/OOXML parsing
stack in their minified bundle. The consumer wants to pay in bundle size only for the formats they
actually use.

> Spec note, not a corpus case: this is a package-architecture and bundle-size concern, not an
> assertable spreadsheet behavior. It cannot be expressed through the roundtrip/inspect adapter
> surface; if verified at all it is via a build-time import-graph or bundle-size test, outside the
> behavior corpus. Concrete KB figures are non-durable motivation, not a contract.

## Desired behavior

- **The document model is format-agnostic.** The in-memory workbook/worksheet model (the "core")
  does not reference any serialization format. Building and mutating a workbook in memory pulls in no
  format parser at all.
- **Each format sits behind its own tree-shakeable entry point.** CSV, XLSX/OOXML, and any future
  formats live behind dedicated subpath exports (or per-format read/write functions) rather than a
  single mutating-accessor facade that eagerly references every format. A standard bundler can drop
  the XLSX stack entirely from a CSV-only application, and vice versa.
- **The package declares itself side-effect-free** (`"sideEffects": false`) so bundlers can eliminate
  unused branches, and the durable requirement is *"unused format code is eliminable by a standard
  bundler,"* verified by an import-graph / bundle test rather than a fixed KB number.

## Prior art / root cause

The reference implementation bundled all parsing behind a single workbook facade
(`workbook.csv.load` / `workbook.xlsx.load` style accessors), coupling every format into one import
graph so a CSV-only or core-only consumer could not shed the XLSX weight. For a TypeScript-first ESM
rewrite this is naturally addressed by (a) a format-agnostic core module, (b) per-format read/write
functions or subpath exports instead of mutating-accessor facades, and (c) `sideEffects: false`.

## Open questions

- Ergonomics versus tree-shakeability: keep a convenience umbrella export (all formats) alongside the
  granular subpath exports — accepting that the umbrella defeats tree-shaking — or force explicit
  per-format imports as the only path?
- Static subpath exports (simpler, SSR-safe, bundler-driven) versus lazy `dynamic import()` code-
  splitting (further browser wins, but async and more complex API)?
- Where does the streaming reader/writer live relative to the core-vs-format split, since streaming
  is XLSX-specific?
- What is the durable verification — an import-graph assertion that the CSV entry point does not
  reach the XLSX modules — rather than a brittle absolute size budget?

Related: `esm-package-entrypoint-ergonomics`, `minimal-audit-clean-dependency-tree`,
`no-global-polyfill-in-browser-bundle`, `unified-streaming-and-buffered-io`,
`published-types-resolve-across-consumers`.
