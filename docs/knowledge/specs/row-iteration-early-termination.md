# Row iteration must support terminating early without throwing

Cluster: streaming

## Scenario

A developer iterates a worksheet's rows to validate uploaded data. On the first invalid row they
want to stop immediately and report the error, not visit every remaining row. With a `forEach`-style
callback iteratee, the legacy shape, there is no supported early break: returning from the callback
ends only that one invocation. The workaround people fall back on is throwing an exception and
catching it around the whole loop, which conflates control flow with real errors and, in async
contexts, risks unhandled rejections.

> Spec note, not a corpus case: this is an API-ergonomics gap with no malformed output to assert,
> since the current iteration produces correct rows and just cannot be stopped cleanly. The durable
> value is the desired iteration API and its stop semantics.

## Desired behavior

- The row and cell iteration API offers a first-class way to terminate early, so
  validate-and-bail patterns are ergonomic and never require exceptions for control flow.
- **Prefer real (async) iterables.** Exposing rows as `Iterable` and `AsyncIterable` lets standard
  `for (const row of ...) { if (bad) break; }` and `for await (const row of ...)` work naturally, so
  `break` and early `return` compose with the enclosing function, and async iteration covers
  streamed reads of large files.
- **If a callback iteratee is kept**, give it explicit stop semantics, for example returning `false`
  halts iteration in the style of `Array.prototype.some` and `every`, and document it. Silent
  "return does nothing" is the trap to remove.
- Early-termination behavior is **consistent** between the in-memory worksheet model and the
  streaming reader.

## Open questions

- Keep a callback iteratee at all, or standardize solely on (async) iterators?
- If both exist, what is the stop signal for the callback form: return `false`, or a passed-in stop
  handle?
- On the streaming reader, can breaking out tear the underlying stream down promptly, without
  leaking file handles or leaving the parser mid-state?
- Does `includeEmpty`, which skips empty rows, change the row numbers and indices exposed to the
  consumer during iteration?

Related: `streaming-read-emits-all-worksheets`, `row-iteration-includes-trailing-empty-cells`,
`streaming-writer-row-commit-backpressure`.
