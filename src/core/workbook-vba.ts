// The workbook's macro project, lifted off `Workbook` into the slice it actually is.
//
// Everything here is reachable from one thing outside itself: the workbook's preserved-reference
// list, where a macro-enabled package's `vbaProject.bin` and its signature closure live. That single
// dependency is why this comes out whole — the list is handed in, and the workbook keeps the public
// accessors that delegate here, exactly as `Worksheet` keeps the accessors in front of its
// validation, conditional-formatting and grid-edit overlays.
//
// The doc comments for the public surface stay on `Workbook`'s accessors, which is what the API
// reference is generated from and what a consumer reads. What is written here is what the
// implementation needs said.

import {
  addVbaReference,
  parseVbaProject,
  removeVbaModule,
  VBA_PROJECT_CONTENT_TYPE,
  VBA_PROJECT_PART_PATH,
  VBA_PROJECT_REL_TYPE,
  VbaAuthorError,
  type VbaLibraryReference,
  type VbaProject,
  type VbaProjectSignature,
  vbaProjectSignatureKind,
} from '../vba/index.ts';
import {replaceContents} from './containers.ts';
import type {PreservedPart} from './preserved.ts';
import type {PreservedWorkbookReference} from './workbook.ts';

/**
 * The macro-project slice of a workbook: the lazily-decoded project, the raw blob the writer
 * re-embeds, the signatures wired off it, and the two structural edits that can be made to a project
 * without recompiling p-code.
 */
export class WorkbookVbaProject {
  // The workbook's own list, held by reference: an attach or a replace has to be visible to the
  // writer, which reads it off the workbook.
  readonly #references: PreservedWorkbookReference[];

  // `#parsed` distinguishes "not yet decoded" from a genuine "no macros" (`undefined`) result, so a
  // macro-free workbook is not re-probed on every access.
  #parsed = false;
  #project: VbaProject | undefined = undefined;

  constructor(references: PreservedWorkbookReference[]) {
    this.#references = references;
  }

  get project(): VbaProject | undefined {
    if (!this.#parsed) {
      const bytes = this.#entry()?.bytes;
      this.#project = bytes ? parseVbaProject(bytes) : undefined;
      this.#parsed = true;
    }
    return this.#project;
  }

  get bytes(): Uint8Array | undefined {
    return this.#entry()?.bytes.slice();
  }

  set bytes(bytes: Uint8Array | undefined) {
    // Validate before touching any state: a malformed blob must fail closed and leave the existing
    // project intact, never half-remove it. Only past this point do we mutate.
    if (bytes !== undefined) parseVbaProject(bytes);

    // Drop any existing project; its whole closure goes, taking a now-stale signature part with it. A
    // fresh reference then mirrors exactly what the reader captures for a macro workbook, so the writer
    // emits a byte-identical macro-enabled package with no writer changes.
    replaceContents(
      this.#references,
      this.#references.filter((r) => !r.relType.endsWith('/vbaProject')),
    );
    if (bytes !== undefined) {
      this.#references.push({
        relType: VBA_PROJECT_REL_TYPE,
        entryPath: VBA_PROJECT_PART_PATH,
        parts: [
          {
            path: VBA_PROJECT_PART_PATH,
            contentType: VBA_PROJECT_CONTENT_TYPE,
            bytes: bytes.slice(),
            rels: [],
          },
        ],
      });
    }
    this.#parsed = false;
    this.#project = undefined;
  }

  // Walk the VBA project's preserved closure for its signature parts — each reached by a signature
  // relationship off `vbaProject.bin`. Computed on each access rather than memoised: the closure is
  // small and already in memory, and recomputing sidesteps a cache that a signature-dropping mutation
  // (bytes replace, module remove, reference add) would otherwise have to invalidate.
  get signatures(): readonly VbaProjectSignature[] {
    const ref = this.#ref();
    const entry = ref?.parts.find((p) => p.path === ref.entryPath);
    if (ref === undefined || entry === undefined) return [];
    const partByPath = new Map(ref.parts.map((p) => [p.path, p]));
    const signatures: VbaProjectSignature[] = [];
    for (const rel of entry.rels) {
      const kind = vbaProjectSignatureKind(rel.type);
      const part = kind === undefined ? undefined : partByPath.get(rel.targetPath);
      if (kind !== undefined && part !== undefined) {
        signatures.push({kind, bytes: part.bytes.slice()});
      }
    }
    return signatures;
  }

  // Both edits go out through `bytes`, so each one re-validates the spliced blob and drops the stale
  // signature the same way an outright replacement does.
  removeModule(name: string): void {
    const bytes = this.bytes;
    if (bytes === undefined) {
      throw new VbaAuthorError('workbook has no VBA project to remove a module from');
    }
    this.bytes = removeVbaModule(bytes, name);
  }

  addReference(ref: VbaLibraryReference): void {
    const bytes = this.bytes;
    if (bytes === undefined) {
      throw new VbaAuthorError('workbook has no VBA project to add a reference to');
    }
    this.bytes = addVbaReference(bytes, ref);
  }

  #ref(): PreservedWorkbookReference | undefined {
    return this.#references.find((r) => r.relType.endsWith('/vbaProject'));
  }

  #entry(): PreservedPart | undefined {
    const ref = this.#ref();
    return ref?.parts.find((p) => p.path === ref.entryPath);
  }
}
