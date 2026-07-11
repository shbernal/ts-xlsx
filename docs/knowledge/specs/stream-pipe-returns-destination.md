# The library's writable streams must honor the Node stream contract (pipe returns the destination)

Cluster: streaming

## Scenario

A caller drives the streaming writer and pipes its output into another stream, using the standard
Node idiom that chains off the destination:

```js
workbook.stream.pipe(fs.createWriteStream('/tmp/sheet.xlsx'))
  .on('finish', () => console.log('written'));
```

or composes it with `stream.pipeline(source, transform, destination, cb)`. Node's `Readable.pipe(dest)`
is contractually required to **return `dest`** so these forms work. The library's internal stream
buffer (`StreamBuf`) overrides `pipe` but returns `undefined`, so `.pipe(dest).on('finish', …)`
throws ("Cannot read properties of undefined") and `stream.pipeline` mis-wires — the finish handler is
never attached and the write appears to hang or silently drop.

> Spec note, not a corpus case: the defect is conformance to Node's stream contract on an internal
> stream helper, and the corpus is implementation-blind — it does not reach into `lib/utils`. The
> durable value is the requirement itself: every stream object the public API hands back behaves like
> a real Node stream. It becomes a corpus case once the public streaming surface exposes a pipeable
> stream that a case can pipe and assert the return value on.

## Desired behavior

- **Every writable/readable stream the public API exposes conforms to the Node stream contract.** In
  particular `pipe(dest)` returns `dest`, so `writer.stream.pipe(out).on('finish', …)` chains and
  `stream.pipeline(...)` wires source→destination correctly.
- **Backpressure and `finish`/`end`/`error` events propagate** as a consumer expects from a native
  stream, so the streaming writer drops into existing Node stream plumbing without special-casing.
- Prefer building on the platform's stream primitives (or a thin, spec-faithful wrapper) rather than a
  bespoke buffer that re-implements — and diverges from — the stream contract. See
  `public-types-node-stream-portability` for the type-surface half of the same portability goal, and
  `web-streams-io-surface` for the cross-platform stream direction.

## Open questions

- Which internal buffering the streaming writer needs is genuinely custom versus expressible with
  `stream.PassThrough` / `Transform`, whose contract conformance is free.
- Whether the rewrite exposes Node `Readable`/`Writable`, Web Streams, or both at the pipe site (ties
  into `web-streams-io-surface` and `unified-streaming-and-buffered-io`).

Related: `public-types-node-stream-portability`, `web-streams-io-surface`,
`unified-streaming-and-buffered-io`, `streaming-writer-incremental-http-delivery`.
