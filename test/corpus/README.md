# Regression corpus

The corpus is **the product's spine** (see [`../../STRATEGY.md`](../../STRATEGY.md)).
It encodes "correct behavior" as a set of implementation-blind cases that run against
*any* implementation through a thin adapter — so it survives the Phase 3 rewrite and
proves the new code is at least as correct as the old, plus everything the old one got
wrong. A bug without a corpus case is a bug that will return.

## Layout

```
test/corpus/
  cases/*.case.mjs      one harvested behavior cluster, implementation-blind
  adapters/<name>.mjs   binds the contract vocabulary to a concrete implementation
  run.mjs               discovers cases, runs them against an adapter, reports red/green
```

Run it:

```
node test/corpus/run.mjs [--adapter current]
```

## A case

A case module default-exports:

```js
{
  id: 'whole-column-defined-names',              // durable descriptive slug — no number prefix
  cluster: 'address-decoding',
  description: '…',
  provenance: { source: 'upstream-issue' },      // OPTIONAL, disposable trace — never the identity
  behavior: [
    { name, baseline: 'pass' | 'fail', expect(api, assert) { … } },
  ],
}
```

- **`id` / `description`** carry the durable identity: a descriptive slug and the
  *real-world scenario* in prose. Do **not** encode upstream issue/PR numbers here —
  they go meaningless when we finish leaving that project (`harvest-triage` skill).
- **`provenance`** is optional and disposable — a trace of where a case came from, never
  its identity. The durable text must stand entirely without it.
- **`behavior[]`** — each is one assertion about observable behavior. `expect` receives
  the **adapter** (`api`) and Node's strict `assert`; it throws to fail, returns to pass.
- **`baseline`** records what **today's legacy code** does for this behavior:
  `pass` = green now (a *regression lock*), `fail` = a known-open bug the rewrite must
  fix. This is what lets the corpus be "mostly red where bugs are real" *without* a red
  build.

## The adapter contract

Cases never import `lib/`. They call a small, growing vocabulary of capabilities that
the adapter provides — the adapter is the **only** place that knows how an
implementation is shaped. Current vocabulary:

| Capability | Meaning |
|---|---|
| `decodeAddress(ref)` | Decode a single cell/row/column reference → `{col, row, …}` (absent axis = `undefined`). |
| `decodeRange(ref)` | Decode a range reference → corners + serialized dimensions. |
| `probeCellFonts({apply, read})` | On a fresh worksheet, assign a font to each `apply` cell, then return `{ <address>: font }` for the `read` cells — for asserting per-cell style stays local. |
| `roundtripWorkbook(spec)` | Build a workbook from a declarative `spec`, write it to a buffer, read it back, and return a normalized JSON model (`{properties, sheets}`) — for asserting content survives write→read. |
| `inspectPackage(spec)` | Build + write a `spec`, unzip the package, and return raw OOXML-part facts (worksheet-declaration consistency, `pageMargins`, `sheetViews`, table XML, per-cell formula text, well-formedness) — for asserting on what is actually serialized. |
| `tryWriteWorkbook(spec)` | Build + attempt to write a `spec`; return `{ok, error, survivingCells, …}` — for asserting pathological input neither throws nor drops sibling cells. |

The `spec` shape consumed by the three workbook capabilities is documented at the top of
`adapters/workbook-io.mjs` (worksheets with cells, columns, rows, page margins, tables).

Add capabilities only as cases demand them, and add them to **every** adapter. When the
rewrite lands, a `rewrite.mjs` adapter binds the same vocabulary to the new code and
every existing case runs unchanged — the corpus does not move, the implementation does.

## What the runner does with baselines

| baseline | actual | status | fails build? |
|---|---|---|---|
| pass | pass | `✓` green | no |
| fail | fail | `○` known-open | no — this is the corpus's job on the frozen tree |
| pass | fail | `✗` regression | **yes** (exit 1) |
| fail | pass | `↑` newly-fixed | no — but flip the baseline to `pass` |

The rewrite's finish line for an area: **every baseline in it flips to `pass`.**
