// Entry barrel for `@shbernal/ts-xlsx/node`: the part of the library that only Node can run.
//
// One thing lives here today, the streaming writer, and it is here for what it imports rather than
// for what it does. It opens a file (`node:fs`) and hands back a `PassThrough` to pipe
// (`node:stream`), and a bundler resolves imports rather than call graphs: while those two modules
// sat on the graph reachable from the package root, every browser build pulled them in and warned
// or failed, whether or not the page ever streamed anything. This barrel is the boundary
// `docs/knowledge/specs/browser-safe-io-boundary.md` asks for: everything else the package
// exports is now reachable in a tab without a bundler alias (ADR 0040).
//
// The rule for what belongs here is the import, not the intent: a module that needs a Node built-in
// is published from this entry, and a module that does not is published from the entry its subject
// belongs to. `scripts/check-browser-safe.ts` is what enforces that, by walking the graph rather
// than by trusting this comment.
//
// `package.json` resolves this subpath to `entries/node-unavailable.ts` under a bundler's `browser`
// condition, so a browser build that imports it links a module which throws by name instead of one
// that cannot be bundled at all.

export {
  type CalcProperties,
  StreamedRow,
  WorkbookStreamWriter,
  type WorkbookStreamWriterOptions,
  WorksheetStreamWriter,
} from '../io/xlsx/write-stream.ts';
