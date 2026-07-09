# The streaming workbook writer must work in the browser over a platform-neutral sink

Cluster: streaming / browser

## Scenario

A web application renders a large dataset (tens of columns, hundreds of thousands of rows) in the
browser and lets the user export to `.xlsx` without a server round-trip. Buffering the whole
workbook in memory before download exhausts the tab's heap and crashes the export. The user needs to
stream the generated package incrementally to a download sink (a WHATWG `WritableStream`, a File
System Access handle, a save-stream shim) so peak memory stays bounded regardless of dataset size —
exactly as incremental sheet writing already allows on Node.

## Desired behavior

- The streaming/incremental writer is usable in a **browser runtime**, not only Node. Writing a
  large workbook keeps peak memory bounded (proportional to a bounded window of buffered rows plus
  compression state, not to total row count).
- Emitted bytes are pushable to a **WHATWG `WritableStream`** so the browser flushes to disk
  progressively; equivalently, the writer hands back an async-iterable / `ReadableStream` of package
  bytes the caller can pipe anywhere. The public API accepts a **platform-neutral byte sink** rather
  than requiring a Node `Stream`/`EventEmitter`.

## Root cause (durable, environment-agnostic)

The historical blocker was that the writing stack detected "is this a stream?" via Node-specific
`instanceof` checks and Node `Buffer` identity, so a browser writable (`WritableStreamDefaultWriter`)
was rejected even though it satisfied the needed contract. The lesson: **sink/stream detection must
be structural** (duck-typed on the methods actually used — `write`/`close`, or the WHATWG
writer/reader shape) and byte handling must not depend on Node `Buffer` identity. Typing must also
admit the browser sink, so callers do not cast a `WritableStreamDefaultWriter` to a Node `Stream`
type.

## Open questions

- Canonical sink shape for the public API: a WHATWG `WritableStream`, an async-iterable/`ReadableStream`
  we return, or an abstract byte-sink interface adapted per platform.
- The zip/deflate layer must run in-browser (WASM or JS deflate) with bounded memory and streamable
  output; confirm the chosen compressor emits incrementally rather than requiring the whole archive
  in memory.
- Backpressure: honor the sink's `desiredSize`/`ready` so a slow disk or download stream throttles
  row production instead of re-buffering unbounded.
- One neutral sink abstraction serving Node and browser, or two thin adapters over a shared core.

Related: `browser-streaming-read-from-seekable-source` (the read face), `streaming-writer-row-commit-backpressure`,
`streaming-writer-incremental-http-delivery`, `browser-safe-io-boundary`, `no-global-polyfill-in-browser-bundle`,
`load-accepts-arraybuffer-and-typed-arrays`.
