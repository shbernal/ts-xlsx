// What `@shbernal/ts-xlsx/node` resolves to when the environment is not Node.
//
// `package.json` points the `browser` condition at this module, so a bundler building for a tab
// links it instead of `entries/node.ts` and never follows an import to `node:fs` or `node:stream`.
// The names are the same and the types are the real ones (the `types` condition is not switched),
// so the substitution is invisible until something is actually constructed, at which point it
// throws by name rather than failing as `undefined is not a function` from inside a minified chunk
// which is the behaviour `docs/knowledge/specs/browser-safe-io-boundary.md` asks for.
//
// `scripts/check-entries.ts` holds this file's value exports equal to `entries/node.ts`'s, so a
// symbol added there cannot silently become a missing import here.

const WHY =
  'the streaming writer opens files and pipes Node streams, which a browser has neither of. ' +
  'Use writeXlsx (or writeXlsxAsync), which produce the same package as bytes.';

function unavailable(name: string): never {
  throw new Error(`ts-xlsx: ${name} is not available in this environment: ${WHY}`);
}

/** Not available outside Node; see {@link WorkbookStreamWriter}. */
export class StreamedRow {
  constructor() {
    unavailable('StreamedRow');
  }
}

/** Not available outside Node; see {@link WorkbookStreamWriter}. */
export class WorksheetStreamWriter {
  constructor() {
    unavailable('WorksheetStreamWriter');
  }
}

/**
 * Not available outside Node. The streaming writer needs `node:fs` and `node:stream`; in a browser,
 * build the workbook in memory and call `writeXlsx`.
 */
export class WorkbookStreamWriter {
  constructor() {
    unavailable('WorkbookStreamWriter');
  }
}
