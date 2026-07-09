# Delete sheet before streaming write

## Remove an uncommitted worksheet from the streaming writer

### Desired behavior
The streaming (incremental) xlsx writer should expose a supported way to remove a worksheet that has been added to the workbook but has **not yet been committed** to the output stream.

- If a worksheet has never had any content committed, removing it should leave the finalized workbook as if that sheet was never added: no sheet entry, no orphaned relationship, no gap in sheet ordering, and correct `sheetId`/index bookkeeping for the remaining sheets.
- Attempting to remove a worksheet that has **already been committed** must fail loudly (streaming output is append-only; you cannot un-write bytes already flushed). The error should clearly state that the sheet was already committed.
- Removal should be addressable the same way sheets are otherwise referenced (by the worksheet handle/name/id), not by manipulating a private array.

### Prior art / why it's not trivial
The streaming writer commits worksheet XML incrementally to keep memory flat, which is why the classic in-memory "remove sheet" operation isn't directly available. The historically-suggested workaround was to splice the entry out of the writer's private internal worksheets collection, which bypasses relationship/content-type/ordering bookkeeping and is unsupported and brittle.

### Open questions
- Should removal be an explicit `removeWorksheet(...)` on the streaming writer, or expressed as the worksheet never being committed (e.g. discard-on-empty)?
- Should the writer optionally auto-drop worksheets that were added but never received any committed rows at finalize time, or only remove on explicit request?
- Interaction with defined names, shared formulas, or cross-sheet references that pointed at the removed sheet — reject the removal, or leave dangling refs to the caller?
