// Entry barrel for `@shbernal/ts-xlsx/xlsx` — the XML codec: whole-workbook read and write, the
// streaming pair, and the in-place VBA part edits.
//
// The streaming reader's entry points are public; the granular per-row/cell/sheet output shapes
// (`StreamedRow`/`StreamedCell`/`StreamedSheet` as *read*) are intentionally left as inferred
// structural types rather than named commitments while that surface settles. The streaming
// *writer*'s whole surface is named: its incremental workbook/worksheet/row handles are classes
// and their options are interfaces, so there is nothing structural left un-named. A styled row
// hands back `Cell`, which belongs to `/core`.
//
// Streaming is not its own entry point. Measured, `read-rows` + `write-stream` reach every module
// `read` + `write` do plus three — an entry that costs what the codec costs is an alias, not a
// packaging boundary.

export {editXlsxVbaAddReference, editXlsxVbaRemoveModule} from '../io/xlsx/edit-vba.ts';
export {type ReadXlsxOptions, readXlsx} from '../io/xlsx/read.ts';
export {
  type ReadSheetRowsOptions,
  readSheetRows,
  readWorkbookStream,
} from '../io/xlsx/read-rows.ts';
export {type WriteOptions, writeXlsx} from '../io/xlsx/write.ts';
export {
  type CalcProperties,
  StreamedRow,
  WorkbookStreamWriter,
  type WorkbookStreamWriterOptions,
  WorksheetStreamWriter,
} from '../io/xlsx/write-stream.ts';
