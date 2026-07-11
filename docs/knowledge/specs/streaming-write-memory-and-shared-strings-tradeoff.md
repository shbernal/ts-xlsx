# High RAM usage on File I/O writing

## Desired behavior

Writing a large workbook must have a predictable, documented, and bounded memory profile. Two paths exist and their memory characteristics must be explicit in the API and docs:

1. **Buffered write** (build the whole model, then serialize): peak memory is proportional to the full workbook model plus the serialized bytes. This is acceptable for small/medium documents but is the wrong tool for large exports. The docs and types should steer users toward streaming for large data.

2. **Streaming write** (row-committed, incremental flush to a sink): peak memory should stay roughly constant with respect to row count, independent of total document size, so that million-row exports finish in tens of MB rather than gigabytes.

## Prior art / observed reality (upstream)

- Buffered writes of ~360k cells produced a ~2.3 MB file but drove RSS to 600 MB–1.5 GB, and the memory was not promptly reclaimed even after nulling references; a forced GC only recovered part of it. This points at retained references / large intermediate buffers, not merely lazy GC.
- The streaming writer dropped peak memory to ~100–400 MB for the same workload, with periodic dips at row commit.
- The dominant remaining memory cost in the streaming path was the **shared-strings table**: with `useSharedStrings: true`, every unique string is hashed and cached for the whole write, so memory grows linearly with distinct string count. Disabling shared strings let a 1M-row export peak around ~55 MB.
- A sharp memory spike was observed at worksheet/workbook **commit** time, where buffered-but-not-yet-flushed content is written out.

## Design implications for the rewrite

- The streaming writer's memory should be bounded by a small window of in-flight rows plus fixed overhead — not by cumulative row count. Shared-strings dedup must not silently make it O(unique strings); if shared strings are supported in streaming mode, either bound the table, make it opt-in with a clear memory warning, or spill it.
- `useSharedStrings` default and its memory implication must be documented at the API surface. Enabling it trades file size for memory; the tradeoff must be discoverable, and the effective default must be stated (not left ambiguous across versions).
- Flushing should be incremental so commit does not require materializing the entire sheet at once; a large end-of-write spike indicates buffered content that should have been streamed to the sink earlier.
- Provide a way to release/dispose a workbook writer so a long-lived server process does not accumulate retained state across successive exports.

## Open questions

- Should shared strings be supported at all in the streaming writer, given the inherent unbounded-cache tension, or only in the buffered writer? Consider a bounded/LRU dedup table as a middle ground.
- What is the target peak-memory guarantee we are willing to assert (e.g. constant + O(columns) + O(bounded-string-window)) and can it be covered by a memory-ceiling test in CI rather than a corpus behavior case?
- Repeated reads (readFile) were also reported to grow arrayBuffer memory monotonically across calls; confirm whether the reader retains buffers and whether that is the same root cause (retained references) or a separate leak worth its own investigation.

## XML-fragment accumulation on very large writes

A distinct symptom of the same underlying inefficiency: generating a workbook with a great deal of
data fails outright with a JavaScript array/range error, not merely high RAM. The XML serializer
accumulates the output as an array of fragments and pushes **three entries per cell** — open tag,
data, close tag — so a wide, tall sheet inflates the intermediate array toward engine limits (max
array length / string length) before it is ever flushed. Emitting each element as a single fragment
(open+data+close combined) roughly thirds the entry count, and — more importantly for the rewrite —
the serializer must **not build the whole document in one in-memory array at all**: it should stream
fragments to the sink incrementally so the intermediate structure is bounded by the in-flight row
window, never by the total cell count. The array-size failure is a hard ceiling that a bounded,
incrementally-flushed writer never approaches. See `bounded-memory-large-workbook-read` for the read
side and `streaming-writer-row-commit-backpressure` for the flush cadence.
