// The convenience entry point: everything the package exports, under one specifier.
//
// It is a union of the subpath entry barrels in `src/entries/`, which are the real public faces:
// `@shbernal/ts-xlsx/core`, `/xlsx`, `/xlsb`, `/csv`, `/vba`, `/customui`, `/errors`. Each symbol
// is listed in exactly one of them, so there is no second list to keep in step here and a star
// re-export cannot silently drop a name to an ambiguity. `scripts/check-entries.ts` is what holds
// the entries disjoint, which is the check that matters here because the failure it prevents is the
// silent one; `scripts/check-layering.ts` separately keeps this file the only composer of them.
//
// One entry is deliberately absent: `/node`, the streaming writer, which imports `node:fs` and
// `node:stream`. Unioning it here would put those on the graph of every consumer who wrote
// `from '@shbernal/ts-xlsx'`, browser builds included, for a symbol they never named, which is
// how they got there in the first place (ADR 0040). Everything reachable from this file runs in a
// tab, and `scripts/check-browser-safe.ts` walks the graph to keep that true.
//
// Importing from here is the right default; it costs nothing extra to a bundler, because
// `sideEffects: false` lets an unused module be dropped whole. Reach for a subpath when the
// consumer has no bundler to do that for it, or when you want the module graph itself to say
// which half of the library a service depends on.

export * from './entries/core.ts';
export * from './entries/csv.ts';
export * from './entries/customui.ts';
export * from './entries/errors.ts';
export * from './entries/vba.ts';
export * from './entries/xlsb.ts';
export * from './entries/xlsx.ts';
