# Streaming write must release committed-row memory per worksheet

## The scenario

A producer streams a large dataset to disk across several worksheets at once — paging
through a data source and appending rows to every sheet on each page, so no sheet can be
committed until all pages are read. Committing a row (`row.commit()`) is supposed to
flush it and free its memory immediately. The point of the streaming writer is to run in
memory proportional to the number of open sheets, not the total rows written.

Today, committed-row memory is released only for the lowest-indexed open worksheet.
Rows committed on any higher-indexed sheet accumulate in memory until that earlier sheet
is committed and closed, so an interleaved multi-sheet write grows without bound and
OOMs on servers writing millions of rows across several sheets. The known community
workaround — write N separate single-sheet workbooks and merge — confirms the constraint
is per-archive, not per-sheet.

## Desired behaviour

- `row.commit()` on any worksheet releases that row's retained memory independently of
  the commit state of every other worksheet.
- An interleaved multi-sheet streaming write runs in memory bounded by the number of
  open sheets (and the current per-sheet backlog), not the cumulative row count.
- The public streaming API makes the memory contract explicit: which calls flush, and
  what stays retained until worksheet commit.

## Root cause (legacy)

Worksheet output is a zip entry stream, and the underlying archive layer consumes
appended entry streams **sequentially** — it does not drain a later entry's stream until
the earlier entries finish. So a higher-index sheet's buffer fills and backpressures but
never flushes to the zip until the earlier sheet commits and its entry stream ends. The
per-row commit logic itself is correct; the retention lives in the archive layer's
serial consumption of entry streams.

## Open questions for the rebuild

- Does the planned zip layer support concurrent/interleaved entry writes? If yes, true
  per-sheet streaming is first-class. If it must serialize entries, the streaming API
  should either (a) document that interleaved multi-sheet writing is bounded only by
  explicitly ordering sheet commits, or (b) spill overflow committed rows for
  not-yet-flushable sheets to a temp file rather than holding them in heap.
- Whether to expose a per-sheet high-water-mark or backpressure signal so a producer can
  pace itself.
- This is the write-side companion to the read-side memory contract in
  [[bounded-memory-large-workbook-read]]; the two should share a documented memory model.
- Not a corpus case: memory release is a resource property, not a JSON-serializable
  observable output — it belongs in the streaming design spec.
