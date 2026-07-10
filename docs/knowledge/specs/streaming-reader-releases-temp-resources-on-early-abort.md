# The streaming reader must release its temporary resources when iteration is abandoned early

Cluster: streaming

## Scenario

A long-running server streams large spreadsheets with the streaming workbook reader, iterating
worksheets (or rows) with `for await`. Very often the caller needs only the first worksheet or the
first few rows and stops early — a `break`, an early `return`, or a thrown error. While parsing, the
streaming reader spills unzipped worksheet parts to temporary files on disk. When iteration runs to
completion the reader cleans those up; when iteration is abandoned before the reader is exhausted,
the cleanup path never runs and the temporary files persist until process exit. On a process that
stays up for weeks across many such requests, the temp directory grows without bound (tens of GB
observed), eventually exhausting free disk — a real availability/DoS-adjacent failure.

> Spec note, not a corpus case: probing confirms the leak is real — an early-broken `for await` over
> the streaming reader leaves extra files in the OS temp directory. But it is a poor corpus fit for
> two reasons: (1) asserting it by counting `os.tmpdir()` entries is flaky (other processes write
> there concurrently), and (2) a corpus case that reproduces the leak would itself pollute the CI
> runner's temp directory on every run — the very unbounded growth it tests for. The durable value is
> the invariant and the precise cleanup mechanism; lock it in the rewrite behind an **injectable temp
> factory / resource tracker** so the assertion observes only the reader's own resources, not the
> whole temp dir.

## Desired behavior

- **Early termination releases the same resources as completion.** Aborting worksheet/row iteration
  via `break`, `return`, or a thrown error must delete every temporary file the reader created and
  close every open handle — leaving the reader's live-resource count back at its pre-parse baseline,
  identical to the run-to-completion case.
- **Use the async-iteration teardown hook.** A `for await` loop that exits early calls `.return()` on
  the iterator; the reader's async generator must release its temporary resources in a `finally` (or
  equivalent teardown) wrapped around its yield points, so no artifact survives an early exit. This is
  the exact mechanism the fix must use, not a best-effort cleanup on a later event.
- **Testability by construction.** The rewrite's streaming reader should accept an injectable
  temp-file factory (or expose a resource tracker), so a regression test can drive an early break /
  throw and assert net-zero live temporary resources without depending on the shared OS temp dir.
- **Bounded by design.** Combined with bounded, streamed decompression (a zip-bomb concern too), the
  reader never accumulates unbounded on-disk or in-memory state across many partial reads.

## Open questions

- Does the rewrite spill to temp files at all, or decompress parts in bounded memory? If it avoids
  temp files, the leak class disappears — but the `.return()`/`finally` teardown discipline still
  applies to any acquired resource (open zip entries, decompression streams).
- Should abandoning iteration be observable to the caller (a `close()`/`dispose()` the caller can call
  explicitly, in addition to the implicit `.return()`), for callers not using `for await`?
- What is the guaranteed timing of cleanup — synchronous with the `.return()` resolution, so a caller
  can rely on the temp files being gone by the time the loop statement completes?

Related: `bounded-memory-large-workbook-read`, `streaming-read-emits-all-worksheets`,
`streaming-read-resolves-shared-strings`, `streaming-reader-styles-option-and-defaults`.
