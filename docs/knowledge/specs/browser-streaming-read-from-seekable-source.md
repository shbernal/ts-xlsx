# Streaming read from a seekable source (browser Blob / random-access bytes)

Cluster: streaming

## Scenario

A browser application lets a user pick a large spreadsheet (hundreds of MB) and must iterate its
rows without loading the whole document into memory. Today the only browser load path buffers the
entire file (and its decompressed contents), exhausting the tab's heap on large files. Because a
browser File/Blob is randomly seekable, a reader can read the zip end-of-central-directory, then
the central directory, then decompress and parse only the entry it currently needs (worksheet
XML, shared strings), streaming rows out as an async iterator — the same contract the Node reader
already offers, but sourced from a Blob rather than a file-descriptor stream.

## Desired behavior

Expose a browser-capable streaming reader that accepts a **seekable source** (a Blob/File in the
browser, or any random-access byte source) and yields worksheets and rows incrementally as an
async iterable, without ever holding the full compressed archive or full decompressed payload in
memory at once. The public shape mirrors the existing streaming-reader contract (async-iterate
worksheets, async-iterate rows within a worksheet, per-row cell access), differing only in the
input type. Memory scales with the largest single zip entry being processed plus a bounded
row/parse buffer, not with total file size.

## Prior art

The Node streaming reader already provides the async-iterator worksheet/row contract over a byte
stream; this extends it to a random-access source. The upstream proposal implemented
seekable-zip reading (read EOCD → central directory → inflate individual entries on demand) so
only the needed entry is decompressed, with materially lower peak memory than buffer-everything.

## Open questions

- **Shared resources.** Shared strings and styles must be read before (or lazily during) row
  iteration. Decide whether the reader eagerly loads them (bounding their size) or resolves them
  lazily, and document the memory tradeoff. See `streaming-read-resolves-shared-strings`.
- **Source abstraction.** Define a small interface (`size` + `read(offset, length)`) that a Blob,
  a File, a Node file handle, or an in-memory buffer can all satisfy — not bound to the DOM Blob
  type. Keeps the path off `browser-safe-io-boundary` violations.
- **Hostile-input hardening (mandatory).** Cap decompressed entry sizes, cap central-directory
  entry count, reject implausible offsets, so a crafted archive cannot force unbounded allocation
  while "streaming".
- **Error semantics.** When a worksheet entry is truncated or the central directory is
  inconsistent, partial iteration must fail loudly rather than silently yield fewer rows.
- Whether browser write-side streaming is in scope, or read-only for now.

Related: `bounded-memory-large-workbook-read`, `streaming-read-resolves-shared-strings`,
`browser-safe-io-boundary`.
