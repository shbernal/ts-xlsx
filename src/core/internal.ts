// The codec's back channel into the model.
//
// A reader has to push state into a `Workbook` that no authoring path can produce: a `<dxfs>` table
// whose indices existing rules already point at, a theme part carried as opaque bytes, a protection
// credential in finished agile form with no recoverable password. The streaming writer likewise has
// to drop a row's cells the moment they are serialised. All of that used to be public methods on the
// model classes — so they shipped in the `.d.ts`, appeared in the generated API reference, and a
// caller who tried one put the workbook in a state nothing else could reach or repair. There was no
// authoring/codec boundary at all: the model class *was* the codec's mutation interface.
//
// These symbols are that boundary. The operations still live on the classes and still mutate exactly
// the state they always did — this moves no data — but reaching them requires a symbol that is not on
// the public barrel and, because `package.json` exposes only the root entry, is not importable from
// outside the package at all. Inside the tree, the import is the audit trail: `grep` for this module
// and you have the complete list of code that can restore preserved state.
//
// Prefer an ordinary public method whenever a caller could reasonably want the operation. This is for
// operations that are *only* meaningful mid-deserialisation.

/**
 * Keys the codec-only operations on `Workbook` and `Worksheet` (see `WorkbookInternals` /
 * `WorksheetInternals`, declared beside their classes).
 */
export const INTERNAL: unique symbol = Symbol('ts-xlsx codec channel');

/**
 * Keys a `Cell`'s link to a named cell style — its OOXML `xfId`. A single hidden property rather than
 * a channel object, so it stays a prototype accessor: cells are the one model class allocated in the
 * millions, and a per-instance internals object would be a real cost for state most cells never carry.
 */
export const NAMED_STYLE_ID: unique symbol = Symbol('ts-xlsx named style link');
