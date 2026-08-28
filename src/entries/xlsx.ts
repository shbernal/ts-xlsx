// Entry barrel for `@shbernal/ts-xlsx/xlsx`, the XML codec: whole-workbook read and write, the
// streaming reader, and the in-place VBA part edits.
//
// The streaming reader's entry points are public; the granular per-row/cell/sheet output shapes
// (`StreamedRow`/`StreamedCell`/`StreamedSheet` as *read*) are intentionally left as inferred
// structural types rather than named commitments while that surface settles.
//
// Streaming is not its own entry point. Measured, `read-rows` + `write-stream` reach every module
// `read` + `write` do plus three, and an entry that costs what the codec costs is an alias, not a
// packaging boundary. The streaming *writer* is nonetheless published from `/node` rather than
// here, and that is not a second opinion about size: it imports `node:fs` and `node:stream`, and an
// entry a browser can resolve may not reach a Node built-in (ADR 0040). Its whole surface is named
// over there; a styled row hands back `Cell`, which belongs to `/core`.

export {editXlsxVbaAddReference, editXlsxVbaRemoveModule} from '../io/xlsx/edit-vba.ts';
export {type ReadXlsxOptions, readXlsx} from '../io/xlsx/read.ts';
export {
  type ReadSheetRowsOptions,
  readSheetRows,
  readWorkbookStream,
} from '../io/xlsx/read-rows.ts';
export {
  DEFAULT_THEME_XML,
  parseThemeColorScheme,
  parseThemeFontScheme,
} from '../io/xlsx/theme-xml.ts';
export {type WriteOptions, writeXlsx, writeXlsxAsync} from '../io/xlsx/write.ts';
