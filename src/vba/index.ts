// Native VBA read and authoring: decode a macro-enabled workbook's `vbaProject.bin` into readable module
// source, and synthesize a valid `.bin` from module source (the §2.3 authoring line — encode primitives
// here, the public surface still landing).
//
// The read path is a projection over bytes the model preserves opaquely on round-trip. See
// docs/plans/vba-read-and-handling.md for the design invariant and the wider VBA feature map.

export {
  type CfbNode,
  type CfbStorage,
  type CfbStream,
  writeCompoundFile,
} from './cfb-writer.ts';
export {VbaAuthorError, VbaParseError} from './errors.ts';
export {compressContainer, decompressContainer} from './ms-ovba.ts';
export {
  parseVbaProject,
  VBA_PROJECT_CONTENT_TYPE,
  VBA_PROJECT_PART_PATH,
  VBA_PROJECT_REL_TYPE,
  type VbaModule,
  type VbaModuleKind,
  type VbaProject,
} from './project.ts';
