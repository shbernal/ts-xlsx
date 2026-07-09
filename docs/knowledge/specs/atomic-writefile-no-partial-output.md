# File writes must be atomic and never leave a zero-byte or partial file

Cluster: xlsx-io / robustness

## Scenario

A caller writes a workbook to disk and observes a zero-byte or truncated file that spreadsheet
applications reject. In the reported cases the symptom traced to three non-library causes: (1) the
write promise was never awaited, so the process moved on — or the file was served/downloaded —
before bytes were flushed; (2) the target directory was read-only in the deployment environment, so
the write silently produced no content; and (3) a batch job fired hundreds-to-thousands of
concurrent write calls in a tight loop with no backpressure, exhausting file handles/memory and
leaving many outputs empty. Users also saw the on-disk file momentarily at zero bytes and only later
at its true size — the write is not atomic, so a consumer reading the path mid-write sees a partial
file.

## Desired behavior

The file-writing API is a Promise that resolves **only after all bytes are durably flushed and the
file is complete**; callers must await it before consuming the path. Two durable product concerns
fall out:

1. **Atomicity.** A write must not expose an observable zero-byte or partially-written file at the
   destination. Prefer writing to a temporary sibling and renaming into place on success, so any
   observer sees either the previous file or the complete new one — never a truncated intermediate.
   On error (read-only directory, disk full) the API rejects with a clear, actionable error rather
   than resolving after producing an empty file.
2. **Bulk / large-volume writing and backpressure.** The buffered write path scales poorly for very
   large output or many files and can appear to hang; the **streaming writer** is the intended path
   for large output. Guidance to document and, where possible, enforce: generating N files must not
   be N un-awaited writes fanned out at once — provide a streaming writer that bounds memory, and
   document that batch generation must serialize or bound concurrency.

## Prior art

The reference implementation exposed both a buffered file-write and a separate stream-based writer;
the buffered path was where users hit zero-byte/hang symptoms under high volume, and the
recommended workaround was the streaming writer. The buffered path did not use temp-file-then-rename,
so a mid-write read saw a growing or zero-byte file.

## Open questions

- Should the default file-write always be atomic (temp + rename), given the extra fsync/rename cost?
- Surface an explicit concurrency-bounded batch-write helper, or is documenting "await and bound
  your own concurrency" sufficient?
- Exact error contract when the target path/directory is not writable.

Related: `streaming-writer-incremental-http-delivery`, `streaming-writer-row-commit-backpressure`,
`streaming-write-per-sheet-memory-release`, `foreign-file-read-modify-write-preserves-validity`.
