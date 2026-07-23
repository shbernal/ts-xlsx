// Native VBA read: decode a macro-enabled workbook's `vbaProject.bin` into readable module source.
//
// This is a read-only projection over bytes the model already preserves opaquely on round-trip; it
// never re-serialises them (Workbook keeps emitting the original blob byte-for-byte). See
// docs/plans/vba-read-and-handling.md for the design invariant and the wider VBA feature map.

export {VbaParseError} from './errors.ts';
export {parseVbaProject, type VbaModule, type VbaModuleKind, type VbaProject} from './project.ts';
