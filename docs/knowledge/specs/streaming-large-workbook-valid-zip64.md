# The streaming writer must emit a valid ZIP (ZIP64 when large) and reject over-limit row counts

Cluster: streaming

## Scenario

A user streams a very large worksheet to disk with the streaming workbook writer — many rows, each
with a wide string cell, committing per row, then the worksheet, then the writer. The resulting
`.xlsx` is structurally corrupted: Excel reports "Repair required", and re-reading it with the
library's own streaming reader throws `Error: invalid signature: 0x…` while walking the ZIP, with the
bogus signature value varying by data. A varying signature points at the ZIP container — wrong
central-directory / end-of-central-directory offsets, or missing ZIP64 records once the archive or an
entry crosses the classic 4 GB / 0xFFFFFFFF offset boundary — not at the spreadsheet XML. It
reproduces across every version tried, so it is a long-standing container defect, not a regression.

> Spec note, not a corpus case: the defect only manifests once the streamed output crosses a ZIP64
> size/offset boundary (the reporter used ten million rows), which is far too large to write in CI
> without OOM/timeout. A hostile-or-huge repro belongs in a spec note; a corpus case would stall the
> runner. Moderate streamed workbooks already produce valid, re-readable ZIPs (locked by
> `streaming-writer-produces-valid-zip-package`), so the assertable target here is the large-scale
> container correctness, verified by design review and a bounded large-ish smoke test off the CI path.

## Desired behavior

- **The streamed archive is always a structurally valid ZIP.** Its central directory and
  end-of-central-directory records point at correct entry offsets, so any conformant ZIP reader walks
  every entry without hitting an invalid local-file-header signature — regardless of how large the
  streamed content grows.
- **ZIP64 is emitted when needed.** Once the archive, an entry's compressed/uncompressed size, an
  entry offset, or the entry count crosses the classic 32-bit limits, the writer emits the ZIP64
  end-of-central-directory record + locator and the per-entry ZIP64 extra fields, so offsets/sizes
  beyond 0xFFFFFFFF are represented correctly rather than truncated into a bogus offset.
- **A large streamed workbook is re-readable by the streaming reader**: iterating its worksheets
  succeeds and yields the written rows, not an invalid-signature error.
- **Over-limit row counts fail loudly at write time.** When the requested rows exceed the format
  maximum (1,048,576 per worksheet), the writer raises a clear, typed error when the limit is crossed
  rather than silently emitting a corrupt file. (This is a distinct guard from the ZIP64 fix: even a
  valid ZIP holding an over-tall sheet is not a valid spreadsheet.)

## Open questions

- Whether the underlying zip layer (today's dependency, or the planned lean-zip rewrite) emits ZIP64
  correctly, or whether ZIP64 support is the actual gap — ties directly to `lean-zip-container-strategy`.
- What a meaningful large-ish smoke test looks like (big enough to exercise the >4 GB / >0xFFFFFFFF
  offset path, or at least the entry-count and streaming-CRC paths) without OOM/timeout on CI —
  possibly gated behind an opt-in "slow" tag rather than the default corpus run.
- Whether to expose a streaming-time progress/size signal so a caller can detect approaching limits,
  and how the row-count guard interacts with multi-sheet streamed workbooks.

Related: `streaming-writer-produces-valid-zip-package`, `lean-zip-container-strategy`,
`bounded-recursion-no-stack-overflow`, `streaming-write-memory-and-shared-strings-tradeoff`.
