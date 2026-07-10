# CSV write must stream into a sink with incremental flush and backpressure

Cluster: csv

## Scenario

A caller exports a large dataset to CSV and pipes the output to a slow sink — an HTTP response, a
cloud-storage upload, a file on disk. They expect memory to stay bounded regardless of row count: the
writer should emit rows incrementally and respect backpressure from the sink, not build the whole CSV
in memory and hand back one giant buffer. The reported failure ("CSV write stream not writing with a
large dataset") traced to a caller-side ordering mistake — resolving the write before attaching the
downstream uploader, plus an empty storage connection string — but underneath it is a real product
question: the streaming CSV path must be unambiguous, hard to misuse, and genuinely incremental.

> Spec note, not a corpus case: the specific report is user error (a race between resolving the write
> and wiring the sink), with no library-bug reproduction. The durable value is the CSV streaming
> contract and the API shape that makes correct use the easy path — distinct from the xlsx streaming
> writer's row-commit backpressure.

## Desired behavior

- **Bounded memory, incremental flush.** Writing N rows to CSV holds working memory proportional to a
  small bounded window, not to N. Rows are serialized and flushed to the sink as they are produced;
  the writer never materializes the entire CSV as one buffer for a large dataset.
- **Backpressure is honored.** When the sink is slow (its write returns false / its buffer fills), the
  CSV producer pauses until the sink drains, rather than buffering unboundedly ahead of it.
- **An unambiguous, misuse-resistant API.** It must be clear (1) how to obtain the CSV byte stream —
  return a readable the caller pipes, or accept a writable and resolve only after the final flush; (2)
  that the returned promise resolves **after** all bytes are flushed to the sink, so a caller cannot
  race completion against wiring the downstream consumer; (3) that unrelated options are not silently
  ignored. The correct streaming path should be the obvious one, so the common "resolved before the
  upload was attached" mistake is not expressible.
- **Encoding stays correct under streaming.** The incremental path emits the same UTF-8 bytes as the
  buffered path (see `csv-write-honors-requested-encoding`), with no multibyte character split across
  a flush boundary.

## Open questions

- Return-a-readable vs. accept-a-writable: which is the canonical CSV streaming API (or both, with one
  sugaring the other)? A returned readable composes with `pipe`/`pipeline` and Web Streams; an
  accept-a-writable form resolves-after-flush and is harder to race.
- How does this share machinery with the xlsx streaming writer's backpressure
  (`streaming-writer-row-commit-backpressure`, `streaming-writer-incremental-http-delivery`) — one
  streaming core, two serializers?
- Browser parity: the same bounded-memory CSV streaming over Web Streams for a browser build
  (`web-streams-io-surface`).

Related: `streaming-writer-row-commit-backpressure`, `streaming-writer-incremental-http-delivery`,
`atomic-writefile-no-partial-output`, `csv-write-honors-requested-encoding`, `web-streams-io-surface`.
