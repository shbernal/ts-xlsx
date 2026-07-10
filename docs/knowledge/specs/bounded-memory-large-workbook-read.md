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
- The **workbook-level defined-names table** is a second, independent culprit, distinct from the visible grid. A real financial workbook with a trivial sheet (~80 KB, dimensions A1:AH258) but a ~2.85 MB defined-names block — ~35,000 entries, most of them `#REF!`, thousands large array literals up to ~5 KB each, plus one external link with ~90 cached sheet names — exhausts a 900 MB heap and never resolves. The blow-up happens during *model assembly* (building objects per defined name), not sheet parsing, so its cost tracks the defined-names table, not the worksheet. The parsed model must (a) retain defined names compactly, ideally lazily, without a large object graph per entry, (b) not choke on `#REF!` values or large array-literal values, and (c) round-trip the names and the external link without loss.
- Zip handling matters for the security posture too: decompressing all parts eagerly is both a memory and a zip-bomb concern; bounded, streamed decompression addresses both.
- A distinct hard-failure mode of the same eager path: materializing a single worksheet's
  decompressed XML as **one JavaScript string** throws `RangeError: Invalid string length` once the
  concatenated string would exceed V8's maximum string length (~512 MB) — a valid multi-hundred-MB
  worksheet (a million-plus rows) fails to open at all, not just slowly. The streaming reader must
  decompress and parse such an entry in chunks and never build the full entry as one string, so the
  file opens with every cell read correctly. A naive fixed-boundary chunk split is not a fix: it can
  slice inside a multi-byte UTF-8 sequence or across an XML token and silently corrupt data — chunk
  boundaries must be handled safely (this is the read-side analog of the write-side chunk-boundary
  correctness locked by `stream-read-multibyte-utf8-chunk-boundary`). Because a faithful repro is
  hundreds of MB, this stays a spec/perf-harness requirement, never a corpus fixture (it would OOM
  or stall CI — the same rule as the whole-column-validation memory note).

### Open questions
- What multiple-of-input-size is the target ceiling for the eager path, and should it be enforced/asserted in a perf regression harness (peak RSS or heapUsed under a fixed cap for a fixture of known size)?
- Should the eager `readFile`-style API refuse or warn above a size threshold and steer callers to the streaming reader, or silently degrade to a streaming strategy internally?
- Can the streaming reader expose a truly constant-memory row iterator (independent of sheet size) as a first-class, documented API rather than an add-on?
