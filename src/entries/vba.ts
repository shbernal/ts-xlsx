// Entry barrel for `@shbernal/ts-xlsx/vba` — the native VBA read view and the structural edits.
//
// A deliberately narrower face than `src/vba/index.ts`, which is the *internal* barrel and also
// carries the CFB writer, the MS-OVBA container primitives and the part-path constants that
// `Workbook` and the codecs need. Those are implementation, not API.

export {
  parseVbaProject,
  type VbaModule,
  type VbaModuleKind,
  type VbaProject,
  type VbaProjectSignature,
  type VbaProjectSignatureKind,
} from '../vba/project.ts';
export {
  addVbaReference,
  removeVbaModule,
  type VbaLibraryReference,
} from '../vba/project-editor.ts';
