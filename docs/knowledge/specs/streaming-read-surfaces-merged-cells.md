# The streaming reader surfaces merged-cell ranges

Cluster: streaming

## Scenario

A user parses a very large workbook (tens of megabytes, too big to buffer) with the streaming
reader, and needs to know which cell ranges are merged — to render the sheet, to read a value once
from a merged master, or to skip covered cells. The buffered reader exposes the merge ranges from the
worksheet model, but the streaming reader does not surface them at all: the `mergeCells` element in
the worksheet XML is not emitted to the consumer, so streaming callers cannot recover merge geometry
without abandoning streaming and loading the whole file — defeating the reason they streamed.

> Spec note, not a corpus case: this is a streaming-reader capability gap and an emission-ordering
> design question (when merges are known relative to the rows they cover), not a malformed-output
> bug on a small file the corpus can exercise. It becomes a corpus case once the streaming reader
> emits merges and a small round-trip asserts the ranges.

## Desired behavior

- **The streaming reader surfaces every merged range** a worksheet declares, as part of the same
  worksheet's stream — either as an event/property on the worksheet reader or attached to the emitted
  worksheet — so a streaming consumer gets the same merge geometry the buffered reader provides.
- **Feature parity is the invariant**: a merged-cell range visible to the buffered reader is visible
  to the streaming reader too; the two I/O strategies do not disagree about the document
  (`unified-streaming-and-buffered-io`).
- **Merge role is derivable while streaming**: a consumer can tell a merged master from a covered
  child cell as rows arrive, complementing `cell-merge-role-introspection` on the buffered side.
- **Ordering is defined**: because `mergeCells` follows the sheet data in the XML, the reader states
  clearly whether merges are available before, during, or only after row emission — so a consumer
  streaming row-by-row knows when it can trust the merge set (buffer the ranges and apply at sheet
  end if they arrive last).

## Open questions

- Surface shape: a `worksheetReader.on('mergeCells', …)` event, a `merges` array populated by
  stream end, or merge info attached to each emitted row. Row-attached is friendliest but requires
  the reader to have parsed `mergeCells` before the rows it covers.
- If `mergeCells` genuinely arrives after the rows, is a one-pass API honest, or must the consumer
  accept that merges are only complete at sheet end? Document the guarantee rather than pretend.
- Memory: the merge set is small relative to the sheet, so buffering all ranges is acceptable even in
  a bounded-memory stream; confirm no pathological file inflates it.

Related: `unified-streaming-and-buffered-io`, `cell-merge-role-introspection`,
`streaming-read-emits-all-worksheets`.
