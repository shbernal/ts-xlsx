# Streaming CSV reader: rows as an async iterable

Cluster: streaming

## Scenario

A user needs to process a very large CSV file that does not fit comfortably in memory. The library
already offers a streaming reader for spreadsheet packages that exposes rows as an async iterable,
letting the caller consume rows one at a time without materializing the whole document. The same
capability is wanted for CSV input: open a CSV file (or stream) and iterate its parsed rows
asynchronously, so arbitrarily large CSV files can be read with bounded memory. Today CSV reading is
bulk-only — the whole file is parsed into a worksheet — so a large CSV cannot be consumed
incrementally.

> Spec note, not a corpus case: this is a new capability (an async-iterable streaming CSV row reader),
> not a bug with a reproduction. There is no failing behavior to baseline today; the durable value is
> the desired API surface and its semantics, and the recorded gap (bulk CSV read exists; bounded-memory
> streaming CSV read does not, while streaming already exists for the package format).

## Desired behavior

- **Rows as an async iterable.** A streaming CSV reader yields parsed rows so callers can
  `for await (const row of reader) { … }` over a large CSV file/stream without loading the entire file.
  Each yielded row surfaces its cell values in order (and ideally a row number).

- **Single-level iteration.** The package reader nests worksheet-then-row because a workbook has many
  sheets; CSV has a single logical sheet, so the honest shape is a flat row iterator directly off the
  reader, without a worksheet nesting level.

- **Option parity with the bulk reader.** Parsing options that apply to non-streaming CSV reads —
  delimiter, quote handling, encoding, date parsing/format hints, header handling — apply equally to
  the streaming path, sharing one option vocabulary.

- **Backpressure and early termination.** Iteration respects consumer pace and allows `break` to stop
  reading early, releasing the input without draining the whole file — the bounded-memory promise must
  hold in practice, not just in shape.

- **Modern input sources.** Source at least a file path (convenience) and a Node `Readable`; a Web
  `ReadableStream` is the forward-looking contract. Pick the modern stream surface rather than
  callback-driven parsing.

## Open questions

- API shape: mirror the package reader's two-level iteration for symmetry, or expose a single-level row
  iterator since CSV is inherently one sheet? The single-level iterator is the simpler, more honest
  surface.
- Row shape: yield raw value arrays, or richer row objects with typing/coercion (numbers, dates)
  consistent with the non-streaming reader? Consistency with the bulk reader argues for the latter.
- Header rows: an option to treat the first row as a header and expose subsequent rows keyed by column
  name.
- How this unifies with the broader streaming/buffered IO surface so CSV and package streaming share
  one consumption idiom (see `unified-streaming-and-buffered-io`).

Related: `streaming-read-emits-all-worksheets`, `csv-write-to-stream-incremental-flush`,
`csv-write-sheet-selection`, `unified-streaming-and-buffered-io`, `native-iteration-protocol`,
`web-streams-io-surface`.
