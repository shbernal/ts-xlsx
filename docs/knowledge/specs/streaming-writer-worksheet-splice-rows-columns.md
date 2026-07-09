# worksheet.spliceRows is not a function error when using Streaming XLSX Writer

## Streaming writer: row/column splice on the worksheet

### Problem
The streaming (incremental, memory-bounded) worksheet exposes a narrower editing surface than the in-memory worksheet. In particular, splice-style row and column operations (insert/remove N rows or columns at an index, optionally replacing with new values) are absent on the streaming worksheet. Callers who reuse worksheet code written against the in-memory model, or who simply expect API parity, get a hard missing-method failure with no guidance.

This has been a recurring pain point across many library versions; the community workaround is to copy the in-memory worksheet's splice implementations onto the streaming worksheet object, which ignores the streaming lifecycle (rows already flushed to disk cannot be spliced) and can silently corrupt output.

### Why it is subtle
A streaming writer commits rows to the underlying stream as it goes. Once a row has been serialized and its buffer flushed, it is no longer addressable in memory, so a general "splice at arbitrary index" cannot be honored the way it is for a fully in-memory sheet. Any solution must be explicit about what window of rows is still mutable.

### Desired behavior (options to decide during the rewrite)
1. **Splice over the un-committed buffer.** Support splice/insert/remove only for rows and columns that have not yet been flushed. If the caller targets an index that has already been committed, fail fast with a precise, documented error explaining that the target is beyond the still-mutable window — never a generic "not a function".
2. **Explicit no-support with a good error.** If splice is deliberately out of scope for the streaming writer, still define the method so it exists on the type, and have it throw a clear, actionable error naming the streaming constraint and pointing at the in-memory workbook for random-access editing.
3. **Column splice parity.** Whatever is decided for rows applies equivalently to columns (column splice hits the same gap).

### Requirements the design must satisfy
- The streaming worksheet's public type must not advertise a method it does not implement, and must not silently diverge from the in-memory worksheet's type without an intentional, documented reason.
- The failure mode for an unsupported operation must be a typed, descriptive error, not an undefined-property/TypeError.
- Column splice and row splice are handled consistently.

### Prior art
- The in-memory worksheet already implements row and column splice semantics (index, deleteCount, insert values); use it as the reference for the mutable-window case.
- Community monkey-patch of the in-memory splice onto the streaming worksheet demonstrates demand but is unsafe because it ignores already-flushed rows.

### Open questions
- Does the intended streaming writer keep any rolling in-memory window of recent rows that splice could legitimately operate on, or is every committed row immediately unreachable?
- Should attempting to splice a committed index be a recoverable error or a fatal one that invalidates the stream?
- Is API parity with the in-memory worksheet a stated goal for the streaming writer, or is a deliberately reduced, streaming-appropriate surface acceptable?
