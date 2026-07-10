# Read and write workbooks over WHATWG Web Streams (edge/serverless runtimes)

Cluster: streaming

## Scenario

A developer on a modern non-Node runtime — a service worker, or an edge/serverless host such as
Deno, Cloudflare Workers, or Bun — wants to read a workbook from, and write one to, a WHATWG
`ReadableStream`/`WritableStream`/`TransformStream`. These environments expose the standard Streams
API natively but may not provide Node's `node:stream` module, so an I/O surface built on Node
`Readable`/`Writable` forces users to hand-roll shims that implement just enough of the Node stream
contract to fool the library. The browser **write** side of this is captured separately
(`browser-streaming-workbook-write`); this note is about the **read** side and the general
edge-runtime generalization of both directions.

> Spec note, not a corpus case: this is an I/O-surface design goal spanning runtimes the corpus does
> not exercise. The durable value is the platform-neutral stream contract and its open questions.

## Desired behavior

- The streaming **reader** accepts a WHATWG `ReadableStream` of package bytes directly as its
  source, with no Node `Readable` shim — so an incoming `fetch` body or a File System Access read
  stream can be parsed incrementally on Deno/Workers/Bun/browser.
- The streaming **writer** can target a WHATWG `WritableStream`, or hand back a `ReadableStream` /
  async-iterable of package bytes the caller pipes anywhere.
- The public API is expressed against **platform-neutral byte sources/sinks** (Web Streams,
  async-iterables of `Uint8Array`), confining Node `Readable`/`Writable` to clearly Node-only entry
  points rather than the cross-platform surface.
- Peak memory stays bounded on both directions regardless of workbook size, matching the Node
  streaming guarantees.

## Open questions

- Is the neutral contract "async-iterable of `Uint8Array`" (adapt Web Streams and Node streams to
  it at the edges), or first-class Web Streams with a thin Node adapter — or both?
- Backpressure: how is the consumer's/producer's pull rate honored across the Web Streams boundary
  without buffering the whole package?
- Do we ship distinct browser/edge and Node entry points so neither surface references the other's
  stream types, or one surface that feature-detects?
- Cancellation: a consumer cancelling a `ReadableStream` read must tear the parser down promptly
  without leaking resources (shared with `row-iteration-early-termination`).

Related: `browser-streaming-workbook-write`, `browser-streaming-read-from-seekable-source`,
`public-types-node-stream-portability`, `load-accepts-arraybuffer-and-typed-arrays`.
