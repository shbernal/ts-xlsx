// Native VBA read and authoring: decode a macro-enabled workbook's `vbaProject.bin` into readable module
// source, synthesize a fresh `.bin` from module source, edit an existing module's source in place by
// splicing the original bytes (references and other modules preserved), and add a new module to an
// existing project the same way.
//
// The read path is a projection over bytes the model preserves opaquely on round-trip. For the design
// invariants and the wider VBA feature map see the ADRs: read view (docs/decisions/0016), authoring
// (0017), and edit-existing-source by splice (0018), plus the spec
// docs/knowledge/specs/xlsm-macro-preservation.md.

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
export {
  addVbaModule,
  addVbaReference,
  editVbaModuleSources,
  type VbaLibraryReference,
} from './project-editor.ts';
export {
  type VbaModuleSource,
  type VbaProjectSpec,
  writeVbaProject,
} from './project-writer.ts';
