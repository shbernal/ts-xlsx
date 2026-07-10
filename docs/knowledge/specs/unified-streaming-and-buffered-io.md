# Streaming and buffered I/O should be one implementation, not two

Cluster: streaming

## Scenario

The library grew two largely-separate code paths for reading and writing `.xlsx`: a buffered
(whole-document-in-memory) path and a streaming (forward-only) path. They diverged — the streaming
side implements a strict subset of the buffered side's features, and the two even reach for different
zip libraries. The result is a maintenance and correctness tax: a feature or fix landed on one path
silently misses the other, tests must be written twice, and callers cannot rely on the two paths
producing equivalent documents from equivalent inputs.

> Spec note, not a corpus case: this is an architecture and parity decision for the rewrite, not a
> single reproducible defect. Its durable value is the invariant it sets, which future corpus cases
> (per-feature parity assertions) will enforce.

## Desired behavior

- **One document model, one serializer, two execution strategies.** The OOXML read/write logic lives
  in a single implementation; "streaming" vs "buffered" is a strategy over the *same* code — how much
  is held in memory and when bytes are flushed — not a parallel reimplementation.
- **Feature parity is a correctness invariant.** Any capability available buffered is available
  streaming unless it is *fundamentally* incompatible with forward-only writing (e.g. inserting a row
  above the write frontier). Such exceptions are explicit, few, and expressed in the type surface
  (see `streaming-writer-worksheet-type-fidelity`) — not accidental gaps.
- **A single, modern zip layer serves both** and works in Node and the browser, removing the original
  reason the two paths diverged. The zip path is bounded and hostile-input-safe on both strategies
  (no zip-bomb naïveté, streaming decompression with limits).
- **Tests are written once against the shared model** and parameterized over the strategy, so a fix
  cannot land on one path and miss the other.

## Open questions

- The seam between "held in memory" and "flushed": chunk/commit granularity, and how much look-behind
  a streaming writer buffers before it must commit (bounds the insert-above-frontier question).
- Browser constraints on the streaming *reader* (no incremental file handle) — is streaming read
  meaningful there, or buffered-only in the browser?
- Backpressure and cancellation semantics for the streaming strategy against Web Streams vs Node
  streams (ties to `public-types-node-stream-portability`).

Related: `streaming-writer-worksheet-type-fidelity`,
`streaming-write-sheet-protection-before-autofilter`, `public-types-node-stream-portability`,
`load-accepts-arraybuffer-and-typed-arrays`.
