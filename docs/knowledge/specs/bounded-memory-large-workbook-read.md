# NodeJS running out of memory while reading small file

## Bounded memory when reading and writing workbooks

### Problem
The eager full-workbook read path allocates memory grossly disproportionate to input size. A spreadsheet that is single-digit megabytes on disk can drive resident heap into the multiple-gigabyte range, causing V8 out-of-memory aborts on memory-constrained hosts (small cloud instances, containers with tight heap limits) while succeeding on high-RAM developer machines. The same complaint recurs for the write path: workbook objects and their intermediate representations are not released promptly, so long-running services accumulate memory.

### Desired behavior
- **Memory should scale with a sane multiple of input size, not hundreds of times it.** Reading a well-formed workbook of N bytes should have a peak working-set that is a small, bounded multiple of N (plus the size of the materialized cell model), not an unbounded balloon.
- **A streaming read path must exist and be documented as the recommended way to process large files.** A consumer should be able to iterate rows/cells of a large sheet without ever holding the entire decompressed sheet XML (or the entire cell model) in memory at once.
- **Intermediate parse buffers must be released.** Decompressed part contents, shared-strings tables, and XML parser scratch state should not be retained after the corresponding cells are materialized. Writing should likewise free per-sheet buffers as parts are flushed.
- **No pathological per-cell overhead.** The reported stack pointed at per-cell processing during parse; the cost model must avoid retaining large transient objects per cell (e.g. accumulating strings or closures keyed per cell).

### Prior art / notes
- Reporters observed a competing library reading the same ~10 MB file in roughly 500 MB versus 2+ GB, indicating the blow-up is an implementation characteristic, not an intrinsic OOXML cost.
- Shared strings are a common culprit: a large shared-strings part fully materialized as JS strings, plus a parallel index, plus the cell model referencing them, triples the footprint. Deduplication and lazy/interned strings help.
- Zip handling matters for the security posture too: decompressing all parts eagerly is both a memory and a zip-bomb concern; bounded, streamed decompression addresses both.

### Open questions
- What multiple-of-input-size is the target ceiling for the eager path, and should it be enforced/asserted in a perf regression harness (peak RSS or heapUsed under a fixed cap for a fixture of known size)?
- Should the eager `readFile`-style API refuse or warn above a size threshold and steer callers to the streaming reader, or silently degrade to a streaming strategy internally?
- Can the streaming reader expose a truly constant-memory row iterator (independent of sheet size) as a first-class, documented API rather than an add-on?
