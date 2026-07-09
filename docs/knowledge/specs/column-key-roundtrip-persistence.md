# Column keys and a write-then-read round-trip

Cluster: tables

## Scenario

A user defines columns with named keys (e.g. `id`, `name`, `dob`) plus header text and width,
writes the workbook, then later reads it back to update it. In the reopened workbook cells can
still be addressed positionally, but addressing a cell by its former key throws an "invalid
column letter" error and the key association is gone. Users find this surprising because keyed
access is central to how they populate and read row data by name, and they expect keys to
survive persistence the way header text and width do.

## The tension

Column "keys" are a library-level convenience — addressing cells and row properties by a
symbolic name instead of a column index — with **no native representation in the file format**.
The `<col>` elements carry width/hidden/style; header text is just cell content in the first
row; none of them carry a symbolic key. So a naive read of a file the library itself wrote
cannot recover the keys, and callers who relied on keyed access get a hard error after
reopening.

## Options discussed

1. **Reassign on read.** Caller re-sets `columns` with the same key definitions after loading.
   Works but is unergonomic and error-prone; nothing is actually persisted.
2. **Persist keys on the `<col>` element** (or nearby) as extra attributes other consumers
   ignore. Keeps round-trip fidelity for our own files but writes non-standard XML and does not
   survive files written by other tools.
3. **Store keys in a package-level custom part / document metadata.** Survives our own
   round-trip without polluting sheet XML.
4. **Treat keys as strictly ephemeral, in-memory-only.** Never claim they persist; make keyed
   access fail gracefully (or fall back to header-text matching) after a read.

## Open questions

- Do we guarantee key survival for files we authored, for all files, or for none?
- If we persist keys, where (sheet-XML attribute vs. custom package part) and how do we avoid
  breaking foreign-tool compatibility and our own foreign-generator tolerance?
- Should keyed cell access after a read with no key data throw, return undefined, or fall back
  to matching against header-row text?
- What is the typed public-API contract: is `key` documented as write-time-only sugar, or as a
  durable column identity?

Whatever we choose, the contract must be explicit and typed at the API surface so callers are
not surprised, and it must not regress tolerance for files produced by other generators.
