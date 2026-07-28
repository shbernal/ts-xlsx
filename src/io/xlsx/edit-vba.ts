// Structurally edit an existing `.xlsm` at the *package* level: remove a module, or add a reference,
// returning new package bytes with every other part preserved byte-for-byte. Only `xl/vbaProject.bin` is
// rewritten (plus dropping a now-stale signature); worksheets, styles, drawings, and every other part
// ride through untouched.
//
// This is the highest-fidelity way to tweak an existing macro project. The alternative — `readXlsx` →
// the matching `Workbook` method → `writeXlsx` — rebuilds the whole package from the parsed model, which
// re-serialises every part and so only preserves what the model captures. For a rich, real-world
// workbook that round-trip can perturb parts Excel is strict about; splicing the original bytes cannot,
// because it never re-authors anything but the macro project. Use this when the input is a real file
// whose non-macro content must be preserved exactly.
//
// Authoring or editing module SOURCE is not here: Excel runs a module's compiled p-code, not its source,
// so that needs the offline `tools/vba-compiler` (VBIDE), which can produce a whole edited `.xlsm`
// directly (ADR 0019).

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {VbaAuthorError} from '../../vba/errors.ts';
import {
  addVbaReference,
  removeVbaModule,
  type VbaLibraryReference,
} from '../../vba/project-editor.ts';
import {relsPathFor} from '../opc/part-paths.ts';
import {
  parseRelationshipRecords,
  relationshipTargetByType,
  resolveRelativePart,
  resolveWorkbookPart,
} from '../opc/read-opc.ts';

const OFFICE_DOCUMENT_REL = 'officeDocument';
const VBA_PROJECT_REL = 'vbaProject';
// Every signature flavour Excel writes over a VBA project (legacy, agile, V3) shares this local-name
// prefix; all become stale the instant the project's bytes change and must be dropped with it.
const VBA_SIGNATURE_REL_INFIX = 'vbaProjectSignature';

/**
 * Remove a standard module from an existing macro-enabled package's VBA project, returning new package
 * bytes. Every part but `xl/vbaProject.bin` is preserved byte-for-byte (see {@link removeVbaModule} for
 * what changes within it), and any digital signature over the old project is dropped because it cannot
 * validate the new bytes.
 *
 * @throws {VbaAuthorError} if the package carries no VBA project, `name` is not in the project, or names
 *   a `document`/`designer` module.
 * @throws {VbaParseError} if the attached `vbaProject.bin` is malformed.
 */
export function editXlsxVbaRemoveModule(xlsx: Uint8Array, name: string): Uint8Array {
  return applyToVbaProjectPart(xlsx, (bin) => removeVbaModule(bin, name));
}

/**
 * Add a registered (COM type-library) reference to an existing macro-enabled package's VBA project,
 * returning new package bytes. Every part but `xl/vbaProject.bin` is preserved byte-for-byte (see
 * {@link addVbaReference} for what changes within it), and any digital signature over the old project is
 * dropped because it cannot validate the new bytes.
 *
 * @throws {VbaAuthorError} if the package carries no VBA project, or any field of `ref` is invalid (see
 *   {@link VbaLibraryReference}).
 * @throws {VbaParseError} if the attached `vbaProject.bin` is malformed.
 */
export function editXlsxVbaAddReference(xlsx: Uint8Array, ref: VbaLibraryReference): Uint8Array {
  return applyToVbaProjectPart(xlsx, (bin) => addVbaReference(bin, ref));
}

// Shared plumbing for every package-level VBA edit: unzip, locate `xl/vbaProject.bin`, replace it with
// whatever `apply` produces, drop a now-stale signature, and re-zip. `apply` is expected to validate
// fail-closed itself (every project-editor primitive does), so a bad edit throws before `files` is
// touched.
function applyToVbaProjectPart(
  xlsx: Uint8Array,
  apply: (bin: Uint8Array) => Uint8Array,
): Uint8Array {
  // Widen off fflate's `Uint8Array<ArrayBuffer>` element type so spliced/re-serialised parts (whose
  // buffers are `ArrayBufferLike`) assign back into the map.
  const files: Record<string, Uint8Array> = unzipSync(xlsx);

  const binPath = locateVbaProjectPart(files);
  const bin = binPath === undefined ? undefined : files[binPath];
  if (binPath === undefined || bin === undefined) {
    throw new VbaAuthorError('package has no VBA project to edit');
  }

  files[binPath] = apply(bin);
  dropStaleSignature(files, binPath);

  return zipSync(files);
}

// Resolve the package's `xl/vbaProject.bin` part the way the reader does: `_rels/.rels` → the
// officeDocument (workbook) part → its `.rels` → the `vbaProject` relationship, each target resolved
// relative to its referrer. undefined when the package declares no such relationship (a macro-free book).
function locateVbaProjectPart(files: Record<string, Uint8Array>): string | undefined {
  const rootRels = textPart(files, '_rels/.rels');
  if (rootRels === undefined) return undefined;
  const workbookTarget = relationshipTargetByType(rootRels, OFFICE_DOCUMENT_REL);
  if (workbookTarget === undefined) return undefined;
  const workbookPath = resolveRelativePart('', workbookTarget);

  const workbookRels = textPart(files, relsPathFor(workbookPath));
  if (workbookRels === undefined) return undefined;
  const vbaTarget = relationshipTargetByType(workbookRels, VBA_PROJECT_REL);
  if (vbaTarget === undefined) return undefined;
  return resolveVbaTarget(workbookPath, vbaTarget);
}

// The workbook's vbaProject relationship uses a workbook-relative target (`vbaProject.bin`); resolve it
// through the workbook-part rule so both that and an absolute `/xl/vbaProject.bin` land on the part path.
function resolveVbaTarget(workbookPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  if (workbookPath.startsWith('xl/')) return resolveWorkbookPart(target);
  return resolveRelativePart(workbookPath, target);
}

// Editing the project invalidates any signature over it, so remove every signature part the project's
// `.rels` reaches, the relationships that point at them, and their content-type overrides — leaving a
// package that advertises no signature rather than a broken one (mirrors Workbook.vbaProjectBytes).
function dropStaleSignature(files: Record<string, Uint8Array>, binPath: string): void {
  const binRelsPath = relsPathFor(binPath);
  const binRels = textPart(files, binRelsPath);
  if (binRels === undefined) return;

  const signatureRels = parseRelationshipRecords(binRels).filter(
    (rel) => !rel.external && rel.type.includes(VBA_SIGNATURE_REL_INFIX),
  );
  if (signatureRels.length === 0) return;

  let contentTypes = textPart(files, '[Content_Types].xml');
  let rels = binRels;
  for (const rel of signatureRels) {
    const partPath = resolveRelativePart(binPath, rel.target);
    delete files[partPath];
    rels = removeRelationshipById(rels, rel.id);
    if (contentTypes !== undefined)
      contentTypes = removeContentTypeOverride(contentTypes, partPath);
  }

  if (parseRelationshipRecords(rels).length === 0) delete files[binRelsPath];
  else files[binRelsPath] = strToU8(rels);
  if (contentTypes !== undefined) files['[Content_Types].xml'] = strToU8(contentTypes);
}

function textPart(files: Record<string, Uint8Array>, path: string): string | undefined {
  const bytes = files[path];
  return bytes === undefined ? undefined : strFromU8(bytes);
}

// Drop the `<Relationship>` element carrying a given Id, matching the element as a whole (self-closing
// or paired) so attribute order does not matter.
function removeRelationshipById(xml: string, id: string): string {
  return xml.replace(
    /<Relationship\b[^>]*?\/>|<Relationship\b[\s\S]*?<\/Relationship>/g,
    (element) => (new RegExp(`\\bId="${escapeRegExp(id)}"`).test(element) ? '' : element),
  );
}

// Drop the `<Override>` naming a given part path; PartName is the full, unambiguous package path.
function removeContentTypeOverride(xml: string, partPath: string): string {
  return xml.replace(/<Override\b[^>]*?\/>/g, (element) =>
    element.includes(`PartName="/${partPath}"`) ? '' : element,
  );
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
