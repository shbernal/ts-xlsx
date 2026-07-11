# Large-workbook writes must be linear-time and bounded-memory

Cluster: streaming

## Scenario

A user builds a large in-memory workbook — on the order of hundreds of thousands of rows across
roughly twenty columns of plain string/number cells — then serializes it to `.xlsx`. Writing takes
minutes or appears to hang entirely, to the point of being unusable, while comparable libraries write
a similar dataset in tens of seconds. The user has already disabled shared strings and styles to
reduce overhead, yet the non-streaming "materialize the whole worksheet, then serialize" path still
fails to complete in acceptable time and memory.

> Spec note, not a corpus case: this is a performance/resource requirement, not a discrete value
> assertion. Wall-clock speed is machine-dependent and would make a corpus case flaky, and a naive
> repro can effectively hang — neither belongs in the behavior corpus. The requirement is recorded
> here; a bounded-time/bounded-memory regression check belongs in a dedicated perf harness with a
> hard timeout, not in the corpus suite.

## Desired behavior

- **Linear time, bounded memory.** Serializing a workbook to `.xlsx` completes in time roughly linear
  in the number of cells and in predictable, bounded memory. It must never degrade to quadratic time
  or hang. Memory must not scale with total cell count held live.
- **A first-class streaming write path.** For sizes where the eager in-memory model is impractical
  (hundreds of thousands of rows), the library must offer streaming writes that emit rows
  incrementally without ever holding the full serialized worksheet in memory, so throughput stays
  high and memory stays flat. The fork already exposes an incremental sheet-writing capability with a
  shared-strings toggle; this is the intended fast path.
- **Disabling shared strings and styles measurably helps.** For bulk scalar exports, turning off the
  shared-string table and style resolution should reduce both time and memory, and that should be the
  documented recommendation.

## Prior art / root cause

Common performance sinks in an eager write path to watch for: shared-string table lookups that are
O(n) or churn a map, per-row/per-cell object allocation, repeated string concatenation of the sheet
XML instead of streamed chunks, and unnecessary style resolution for plain cells. Competing libraries
reach tens-of-seconds writes at this scale using dense/streamed row emission and cheap
number-to-string paths. Deflate compression level can dominate cost at scale — the default may want
to trade ratio for speed on large sheets.

## Open questions

- Should the ergonomic top-level write API auto-select the streaming path above a size threshold, or
  must the caller opt in explicitly?
- What time and peak-memory envelopes do we commit to (e.g. for 500k rows × 20 columns of scalars),
  and should CI carry a perf-regression guard (time/memory budget) rather than a value assertion?
- Do we guarantee that disabling shared strings and styles measurably reduces both time and memory,
  and document that as the bulk-export recommendation?
- Is deflate level the dominant cost at this scale, and should the large-sheet default favor speed
  over compression ratio?

Related: `streaming-write-memory-and-shared-strings-tradeoff`,
`streaming-write-per-sheet-memory-release`, `bounded-memory-large-workbook-read`,
`unified-streaming-and-buffered-io`, `lean-zip-container-strategy`.
