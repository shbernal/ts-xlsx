# Declaring a table in the streaming workbook writer

Cluster: streaming

## Scenario

A user writes a large spreadsheet through the streaming workbook writer to keep memory bounded, and
wants to define an Excel table over the rows they streamed out — for filtering, sorting, banded
styling, structured references. In the buffered path a worksheet exposes an add-table capability, but
the streamed worksheet either lacks the method or throws at runtime when it is called. Worse, the
worksheet type still advertises the add-table method, so callers discover the gap only at runtime
with a "not a function" error rather than at compile time. Because a table's XML footprint (its
column/header definition and cell range) is known once the row range is committed, declaring a table
in the streaming path without buffering the whole sheet should be possible.

> Spec note, not a corpus case: the feature does not exist yet, so there is no behavior to assert.
> The durable value is the streaming table declaration model and the honest-type requirement.

## Desired behavior

- The streaming writer supports declaring a table over a range of streamed rows, producing the same
  table part XML, header row, autoFilter, table style, and worksheet/table relationship the buffered
  writer produces for an equivalent table.
- **The public type surface is honest**: a streaming worksheet must NOT advertise a table-adding
  method it cannot fulfill. Either (a) the method exists and works in streaming mode, or (b) the
  streaming worksheet type omits it so misuse is a compile-time error — never a silent runtime throw
  against a signature that claims support.
- The table's declared range stays consistent with the actual written extent, the header labels
  match the header cells, and the table never silently overwrites or reflows already-flushed cells
  (the header row and any totals row are part of the streamed content).

## Prior art / constraints

Streamed rows are flushed incrementally and cannot be re-read, so the declaration model differs from
the buffered one. Two workable shapes: (1) declare the table up front (name, header columns, style,
totals intent) then stream the data rows into it, finalizing the cell range on commit from the row
count; or (2) stream rows first, then attach a table over an explicit committed range before
finalizing the sheet.

## Open questions

- Require the header row in the table declaration (authoritative column names), or infer it from the
  first streamed row?
- How are totals rows handled when the row count is known only at commit time?
- When the declared table range and the streamed row range disagree — hard error, or clamp?
- Should the same declarative table spec the buffered round-trip uses also drive the streaming
  writer so both paths share one description?

Related: `streaming-write-add-image`, `streaming-writer-worksheet-splice-rows-columns`,
`table-headerless-omits-autofilter`, `existing-table-roundtrip-fidelity`.
