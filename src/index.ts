// The convenience entry point: everything the package exports, under one specifier.
//
// It is a union of the subpath entry barrels in `src/entries/`, which are the real public faces —
// `@shbernal/ts-xlsx/core`, `/xlsx`, `/xlsb`, `/csv`, `/vba`, `/customui`, `/errors`. Each symbol
// is listed in exactly one of them, so there is no second list to keep in step here and a star
// re-export cannot silently drop a name to an ambiguity. `scripts/check-layering.ts` holds the
// entries disjoint and keeps this file's composition honest.
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
