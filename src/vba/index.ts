// Native VBA read: decode a macro-enabled workbook's `vbaProject.bin` into readable module source.
//
// This is a read-only projection over bytes the model already preserves opaquely on round-trip; it
// never re-serialises them (Workbook keeps emitting the original blob byte-for-byte). See
// docs/plans/vba-read-and-handling.md for the design invariant and the wider VBA feature map.

export {
  type CfbNode,
  type CfbStorage,
  type CfbStream,
  writeCompoundFile,
} from './cfb-writer.ts';
export {VbaAuthorError, VbaParseError} from './errors.ts';
export {
  parseVbaProject,
  VBA_PROJECT_CONTENT_TYPE,
  VBA_PROJECT_PART_PATH,
  VBA_PROJECT_REL_TYPE,
  type VbaModule,
  type VbaModuleKind,
  type VbaProject,
} from './project.ts';
