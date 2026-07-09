# Backlog tracker

The single source of truth for **disposition**: every upstream issue and PR we harvest
is recorded here as `captured | superseded | out-of-scope` with a one-line rationale, so
we can prove nothing was silently dropped (`CLAUDE.md` §"no silent caps"). Raw thread
content lives in [`backlog/issues/`](backlog/README.md); this file is the triage ledger
over it.

- **captured** — distilled into a corpus case and/or spec note (link it).
- **superseded** — obsolete (e.g. a trivial dep bump the rewrite makes moot); say why.
- **out-of-scope** — deliberately not carried; say why.

_Phase 1 has not begun at scale; this file currently holds only the Phase 0 end-to-end
proof. Backlog size at fork time: ~654 open issues + ~139 open PRs._

## Clusters

Seeded from the `STRATEGY.md` snapshot; populated during Phase 1 triage:
`address-decoding` · tables · styles · streaming · pivot · images · conditional-formatting
· dates · formulas · csv · types · security/deps.

## Dispositions

| # | Type | Title | Disposition | Rationale / link |
|---|---|---|---|---|
| [140](https://github.com/exceljs/exceljs/issues/140) | issue | col-cache.js: Cannot read property '0' of null | **captured** | Full-row/column defined-name refs (`$1:$1`, `$A:$A`). Corpus: [`0140-address-decoding`](../../test/corpus/cases/0140-address-decoding.case.mjs). Crash fixed upstream (2 green locks); serialized-address `undefined`/`NaN` leak still open (1 red). Cluster `address-decoding`; see also #134, upstream PR #636. |
