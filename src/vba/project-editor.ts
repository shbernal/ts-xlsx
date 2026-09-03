// Structural edits to an existing `vbaProject.bin` that do NOT touch any module's compiled p-code:
// remove a standard module, or add a registered library reference. Each is a surgical splice: parse the
// original container, rebuild its whole storage/stream tree, and change only the `dir` records (and, for
// a removal, the `PROJECT`/`PROJECTwm` text) the edit demands. Every module stream, `_VBA_PROJECT`, and
// every untouched record rides through byte-for-byte; preservation is by *not touching* them.
//
// Authoring or editing module SOURCE is deliberately NOT here. Excel does not recompile from source on
// open. A module runs the p-code it ships, and only a real Excel can produce genuinely source-matched
// p-code, so source authoring/editing lives in the offline `tools/vba-compiler` (VBIDE), not in this
// pure-TS path (ADR 0019). These splices are safe precisely because they leave every module's p-code
// exactly as its own compiler wrote it.

import {quoted} from '../errors.ts';
import {readU16, spliceBytes} from './bytes.ts';
import {type CfbNode, writeCompoundFile} from './cfb-writer.ts';
import {CompoundFile} from './cfb.ts';
import {type Decoder, decoderForCodePage, type Encoder, encoderForCodePage} from './codepage.ts';
import {
  dirRecords,
  REC_MODULE_NAME,
  REC_MODULE_STREAMNAME,
  REC_MODULE_TERMINATOR,
  REC_MODULES_COUNT,
  REC_REFERENCE_NAME,
  REC_REFERENCE_NAME_UNICODE,
  REC_REFERENCE_REGISTERED,
} from './dir-records.ts';
import {VbaAuthorError, VbaParseError} from './errors.ts';
import {compressContainer, decompressContainer} from './ms-ovba.ts';
import {parseVbaProjectIn} from './project.ts';
import {push, u16, u32, utf16le, validateVbaName} from './vba-encoding.ts';

const DIR_STREAM = 'dir';
const PROJECT_STREAM = 'PROJECT';
const PROJECTWM_STREAM = 'PROJECTwm';
const VBA_STORAGE = 'VBA';

/**
 * Remove a standard module from an existing `vbaProject.bin`, returning new bytes that carry every
 * remaining module, reference, and host-info record unchanged. It drops the module's `VBA/<name>`
 * stream, its MODULE record block in `dir` (decrementing `MODULES_COUNT`), and its `Module=`/`Class=` +
 * workspace lines in `PROJECT`/`PROJECTwm`.
 *
 * Only `procedural` and `class` modules can be removed this way. Removing a `document` module (e.g.
 * `ThisWorkbook`) or a `designer` module (a UserForm) would leave the host referencing code that no
 * longer exists, since their names are tied to a worksheet/workbook `codeName` or a designer storage
 * this project-level primitive has no visibility into. Editing such a module's code-behind is a job for
 * the offline `tools/vba-compiler` (in-place mode), which drives the real host.
 *
 * @throws {VbaParseError} if `bin` is not a parseable VBA project (validated before any edit).
 * @throws {VbaAuthorError} if `name` is not in the project, or names a `document`/`designer` module.
 */
export function removeVbaModule(bin: Uint8Array, name: string): Uint8Array {
  // Parse fail-closed first: validates the container and resolves the module's kind/stream name, so
  // nothing is mutated on a bad input or an unsupported module kind. The container is opened once and
  // shared with the parse, rather than built again below over the same bytes.
  const cfb = new CompoundFile(bin);
  const project = parseVbaProjectIn(cfb);
  const nameKey = name.toUpperCase(); // VBA names are case-insensitive
  const module = project.modules.find((m) => m.name.toUpperCase() === nameKey);
  if (!module) throw new VbaAuthorError(`module ${quoted(name)} is not in the VBA project`);
  if (module.kind !== 'procedural' && module.kind !== 'class') {
    throw new VbaAuthorError(
      `cannot remove module ${quoted(name)}: its kind ${quoted(module.kind)} is tied to host linkage this ` +
        'primitive cannot verify',
    );
  }

  const dirCompressed = cfb.readStream(DIR_STREAM);
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const patchedDir = removeModuleDirRecord(
    decompressContainer(dirCompressed),
    module.streamName,
    project.codePage,
  );

  // Leave _VBA_PROJECT untouched. Resetting it to an "unmatchable version" cookie does NOT force Excel
  // to recompile from source (Excel runs the p-code as-is); on a project that carries real p-code the
  // reset actively crashes the VBA load (verified 2026-07-24, ADR 0019). The surviving modules keep
  // their own compiled p-code; the `dir` stream, authoritative for the module list, no longer names
  // the removed module, which is what makes the removal take.
  const replacements = new Map<string, Uint8Array>([[DIR_STREAM, compressContainer(patchedDir)]]);

  const decoder = decoderForCodePage(project.codePage);
  const encode = encoderForCodePage(project.codePage);
  const projectText = cfb.readStream(PROJECT_STREAM);
  if (projectText) {
    replacements.set(
      PROJECT_STREAM,
      encode(removeProjectStreamLines(decoder.decode(projectText), module.name, module.kind)),
    );
  }
  const projectwm = cfb.readStream(PROJECTWM_STREAM);
  if (projectwm) {
    replacements.set(
      PROJECTWM_STREAM,
      removeProjectwmRecord(projectwm, project.modules.length, module.name, decoder),
    );
  }

  const applied = new Set<string>();
  const withReplacements = replaceStreams(cfb.tree(), replacements, applied);
  if (!applied.has(DIR_STREAM))
    throw new VbaParseError("VBA project 'dir' stream is not in the container tree");

  const removed = new Set<string>();
  const newTree = removeFromStorage(withReplacements, VBA_STORAGE, module.streamName, removed);
  if (!removed.has(VBA_STORAGE)) {
    throw new VbaParseError(
      `module stream ${quoted(module.streamName)} is not in the ${quoted(VBA_STORAGE)} storage`,
    );
  }

  return writeCompoundFile(newTree);
}

/**
 * A registered (COM Automation type-library) reference to add to an existing VBA project: the shape of
 * a real "add a reference to Microsoft Scripting Runtime" call. Project references (to another VBA
 * project) and control references (to an ActiveX control library) are out of scope. See
 * {@link addVbaReference}.
 */
export interface VbaLibraryReference {
  /**
   * The reference's namespace name in the VBA editor: what a qualified reference like
   * `Scripting.Dictionary` resolves through. Must be a valid VBA identifier, at most 31 characters, as
   * real type libraries use (e.g. `Scripting`, `Office`, `stdole`).
   */
  readonly name: string;
  /**
   * The friendly name shown in the References dialog, e.g. `Microsoft Scripting Runtime`. Real projects
   * usually keep this distinct from {@link name}; defaults to {@link name} if omitted.
   */
  readonly displayName?: string;
  /** The type library's GUID, e.g. `{420B2830-E718-11CF-893D-00A0C9054228}` (braces optional). */
  readonly guid: string;
  /** The type library's major version, an integer in `[0, 0xFFFF]` ([MS-OVBA] `LibidMajorVersion`). */
  readonly majorVersion: number;
  /** The type library's minor version, an integer in `[0, 0xFFFF]` ([MS-OVBA] `LibidMinorVersion`). */
  readonly minorVersion: number;
  /**
   * The type library's LCID, an integer in `[0, 0xFFFFFFFF]`. Defaults to `0` (locale-neutral), the
   * overwhelming common case (every reference in a real project observed while building this had `0`).
   */
  readonly lcid?: number;
  /** Absolute Windows path to the type library file, e.g. `C:\Windows\System32\scrrun.dll`. */
  readonly path: string;
}

// LibidMajorVersion/LibidMinorVersion ([MS-OVBA] 2.1.1.8): 1*4HEXDIG, so at most 0xFFFF.
const MAX_LIBID_VERSION = 0xffff;
// LibidLcid: 1*8HEXDIG, so at most 0xFFFFFFFF (practically always 0, locale-neutral).
const MAX_LIBID_LCID = 0xffffffff;
// LibidRegName: *255(%x01-FF), so at most 255 bytes, never NUL.
const MAX_DISPLAY_NAME_CHARS = 255;
const GUID_PATTERN =
  /^\{?([0-9A-Fa-f]{8})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{12})\}?$/;

interface NormalizedReference {
  readonly name: string;
  readonly libid: string;
}

// Validate every field fail-closed and assemble the Libid string ([MS-OVBA] 2.1.1.8 LibidReference ABNF:
// `*\G{GUID}#Major.Minor#LCID#Path#RegName`, hex digit strings with no `0x` prefix), confirmed
// byte-for-byte against a real Excel-authored reference (2026-07-23):
// `*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\System32\scrrun.dll#Microsoft Scripting Runtime`.
function normalizeReference(ref: VbaLibraryReference): NormalizedReference {
  validateVbaName(ref.name, 'reference');

  const guidMatch = GUID_PATTERN.exec(ref.guid.trim());
  if (!guidMatch) throw new VbaAuthorError(`invalid reference GUID ${quoted(ref.guid)}`);
  const guid = `{${guidMatch.slice(1, 6).join('-').toUpperCase()}}`;

  for (const [field, value] of [
    ['majorVersion', ref.majorVersion],
    ['minorVersion', ref.minorVersion],
  ] as const) {
    if (!Number.isInteger(value) || value < 0 || value > MAX_LIBID_VERSION) {
      throw new VbaAuthorError(
        `reference ${field} must be an integer in [0, 0xFFFF], got ${value}`,
      );
    }
  }
  const lcid = ref.lcid ?? 0;
  if (!Number.isInteger(lcid) || lcid < 0 || lcid > MAX_LIBID_LCID) {
    throw new VbaAuthorError(`reference lcid must be an integer in [0, 0xFFFFFFFF], got ${lcid}`);
  }

  if (ref.path.length === 0 || ref.path.includes('\0') || ref.path.includes('#')) {
    throw new VbaAuthorError(
      `invalid reference path ${quoted(ref.path)} (must be non-empty and contain no NUL or '#')`,
    );
  }
  const displayName = ref.displayName ?? ref.name;
  if (
    displayName.length === 0 ||
    displayName.length > MAX_DISPLAY_NAME_CHARS ||
    displayName.includes('\0')
  ) {
    throw new VbaAuthorError(`invalid reference display name ${quoted(displayName)}`);
  }

  const libid =
    `*\\G${guid}#${ref.majorVersion.toString(16).toUpperCase()}.` +
    `${ref.minorVersion.toString(16).toUpperCase()}#${lcid.toString(16).toUpperCase()}` +
    `#${ref.path}#${displayName}`;
  return {name: ref.name, libid};
}

/**
 * Add a registered (COM type-library) reference to an existing `vbaProject.bin`, returning new bytes
 * that carry every existing module, reference, and host-info record unchanged. It grows the project's
 * `dir` stream by one `REFERENCENAME` + `REFERENCEREGISTERED` record pair, positioned immediately before
 * `MODULES_COUNT` (references have no count field of their own; `MODULES_COUNT` simply marks where the
 * reference array ends). It needs no change to `PROJECT`/`PROJECTwm`: a real Excel-authored `PROJECT`
 * stream carries no `Reference=` line at all: references live only in `dir` (confirmed against a genuine
 * Excel-authored project).
 *
 * @throws {VbaParseError} if `bin` is not a parseable VBA project (validated before any edit).
 * @throws {VbaAuthorError} if any field of `ref` is invalid (see {@link VbaLibraryReference}), or the
 *   assembled reference text has a character the project's code page cannot represent.
 */
export function addVbaReference(bin: Uint8Array, ref: VbaLibraryReference): Uint8Array {
  const normalized = normalizeReference(ref);

  // Parse fail-closed first: validates the container before any mutation, over the one container this
  // function opens rather than a second built over the same bytes.
  const cfb = new CompoundFile(bin);
  const encode = encoderForCodePage(parseVbaProjectIn(cfb).codePage);

  const dirCompressed = cfb.readStream(DIR_STREAM);
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const records = buildReferenceDirRecords(normalized, encode);
  const patchedDir = insertReferenceDirRecords(decompressContainer(dirCompressed), records);

  // Leave _VBA_PROJECT untouched; see the note in removeVbaModule. The new reference is unused by the
  // existing modules' p-code, so they load and run unchanged; only the `dir` reference array grows.
  const replacements = new Map<string, Uint8Array>([[DIR_STREAM, compressContainer(patchedDir)]]);

  const applied = new Set<string>();
  const newTree = replaceStreams(cfb.tree(), replacements, applied);
  if (!applied.has(DIR_STREAM))
    throw new VbaParseError("VBA project 'dir' stream is not in the container tree");

  return writeCompoundFile(newTree);
}

// Build the REFERENCENAME + REFERENCEREGISTERED record bytes ([MS-OVBA] 2.3.4.2.2.2 / .2.2.5) for one
// reference. REFERENCENAME's MBCS/Unicode name pair mirrors MODULE_NAME/MODULE_NAME_UNICODE's shape;
// REFERENCEREGISTERED is one record carrying SizeOfLibid + Libid + two zero Reserved fields.
function buildReferenceDirRecords(ref: NormalizedReference, encode: Encoder): number[] {
  const r: number[] = [];
  const nameBytes = [...encode(ref.name)];
  push(r, REC_REFERENCE_NAME, nameBytes);
  push(r, REC_REFERENCE_NAME_UNICODE, utf16le(ref.name));
  const libidBytes = [...encode(ref.libid)];
  push(r, REC_REFERENCE_REGISTERED, [
    ...u32(libidBytes.length),
    ...libidBytes,
    ...u32(0),
    ...u16(0),
  ]);
  return r;
}

// Insert new reference dir records right before MODULES_COUNT (0x000f). The reference array has no
// explicit count field; MODULES_COUNT is simply the next record once the last reference ends (confirmed
// against a real Excel-authored dir stream). Every other record, other references and all modules, rides
// through unchanged.
function insertReferenceDirRecords(dir: Uint8Array, records: readonly number[]): Uint8Array {
  let insertAt = -1;
  for (const {id, recordStart} of dirRecords(dir, 'overruns while adding a reference')) {
    if (id === REC_MODULES_COUNT) {
      insertAt = recordStart;
      break;
    }
  }
  if (insertAt < 0) throw new VbaParseError('dir stream is missing MODULES_COUNT');

  return spliceBytes(dir, insertAt, insertAt, Uint8Array.from(records));
}

// Remove one module's MODULE record block from a decompressed `dir` stream, and decrement MODULES_COUNT.
// A block runs from its MODULE_NAME record (which always opens the block, mirroring buildModuleDirRecord's
// emission order) through its own MODULE_TERMINATOR, identified by matching MODULE_STREAMNAME against
// `streamName`. Every other record (PROJECTREFERENCES, other modules, project-level fields) is carried
// through untouched.
function removeModuleDirRecord(dir: Uint8Array, streamName: string, codePage: number): Uint8Array {
  const decoder = decoderForCodePage(codePage);
  let countAt = -1;
  let blockStart = -1;
  let removeStart = -1;
  let removeEnd = -1;
  let currentStream: string | undefined;
  for (const {id, recordStart, dataStart, size, end} of dirRecords(
    dir,
    'overruns while removing a module',
  )) {
    if (id === REC_MODULES_COUNT) {
      if (size < 2) throw new VbaParseError('PROJECTMODULES MODULES_COUNT record is malformed');
      countAt = dataStart;
    } else if (id === REC_MODULE_NAME) {
      blockStart = recordStart;
    } else if (id === REC_MODULE_STREAMNAME) {
      currentStream = decoder.decode(dir.subarray(dataStart, dataStart + size));
    } else if (id === REC_MODULE_TERMINATOR) {
      if (currentStream === streamName) {
        removeStart = blockStart;
        removeEnd = end;
      }
      currentStream = undefined;
      blockStart = -1;
    }
  }
  if (countAt < 0) throw new VbaParseError('dir stream is missing MODULES_COUNT');
  if (removeStart < 0 || removeEnd < 0) {
    throw new VbaParseError(`module stream ${quoted(streamName)} not found in the dir stream`);
  }

  // MODULES_COUNT always precedes every module block, so countAt is unaffected by removing bytes after it.
  const out = spliceBytes(dir, removeStart, removeEnd);
  const newCount = readU16(out, countAt) - 1;
  out[countAt] = newCount & 0xff;
  out[countAt + 1] = (newCount >> 8) & 0xff;
  return out;
}

// Remove a module's declaration line (`Module=`/`Class=`) and its workspace line from the `PROJECT` text
// stream: the inverse of insertProjectStreamLines. Every other line is left exactly as it was.
function removeProjectStreamLines(
  text: string,
  name: string,
  kind: 'procedural' | 'class',
): string {
  const eol = text.includes('\r\n') ? '\r\n' : '\n';
  const lines = text.split(/\r\n|\r|\n/);
  const declLine = `${kind === 'procedural' ? 'Module' : 'Class'}=${name}`;
  const declIndex = lines.indexOf(declLine);
  if (declIndex >= 0) lines.splice(declIndex, 1);

  const wsIndex = lines.findIndex((l) => l.trim() === '[Workspace]');
  if (wsIndex >= 0) {
    for (let i = wsIndex + 1; i < lines.length; i++) {
      const l = lines[i] as string;
      if (l.trim() === '' || l.startsWith('[')) break;
      if (l.startsWith(`${name}=`)) {
        lines.splice(i, 1);
        break;
      }
    }
  }

  return lines.join(eol);
}

// Remove a module's (MBCS name, UTF-16 name) pair from the binary PROJECTwm stream: the inverse of
// insertProjectwmRecord. `existingModuleCount` (from the already fail-closed-parsed project, before
// removal) bounds the walk to the module records, so it never mistakes the terminator for a record.
function removeProjectwmRecord(
  wm: Uint8Array,
  existingModuleCount: number,
  name: string,
  decoder: Decoder,
): Uint8Array {
  let pos = 0;
  let removeStart = -1;
  let removeEnd = -1;
  for (let i = 0; i < existingModuleCount; i++) {
    const recordStart = pos;
    const mbcsEnd = wm.indexOf(0x00, pos);
    if (mbcsEnd < 0)
      throw new VbaParseError('PROJECTwm record is missing its MBCS name terminator');
    const mbcsName = decoder.decode(wm.subarray(pos, mbcsEnd));
    pos = mbcsEnd + 1;
    let utf16End = pos;
    while (utf16End + 1 < wm.length && (wm[utf16End] !== 0 || wm[utf16End + 1] !== 0))
      utf16End += 2;
    if (utf16End + 1 >= wm.length) {
      throw new VbaParseError('PROJECTwm record is missing its Unicode name terminator');
    }
    pos = utf16End + 2;
    if (mbcsName === name) {
      removeStart = recordStart;
      removeEnd = pos;
    }
  }
  if (removeStart < 0 || removeEnd < 0) {
    throw new VbaParseError(`module ${quoted(name)} not found in the PROJECTwm stream`);
  }

  return spliceBytes(wm, removeStart, removeEnd);
}

// Remove the first direct child stream named `streamName` from the first storage named `storageName`
// found in the tree (depth-first), marking `storageName` in `removed` once done. The inverse of
// insertIntoStorage.
function removeFromStorage(
  nodes: readonly CfbNode[],
  storageName: string,
  streamName: string,
  removed: Set<string>,
): CfbNode[] {
  return nodes.map((n) => {
    if ('data' in n) return n;
    const children = removeFromStorage(n.children, storageName, streamName, removed);
    if (n.name === storageName && !removed.has(storageName)) {
      const filtered = children.filter((c) => !('data' in c && c.name === streamName));
      if (filtered.length !== children.length) removed.add(storageName);
      return {name: n.name, children: filtered};
    }
    return {name: n.name, children};
  });
}

// Rebuild the node tree, swapping any stream whose name has a replacement. Non-stream nodes (storages)
// recurse; everything without a replacement is carried through byte-for-byte.
function replaceStreams(
  nodes: readonly CfbNode[],
  replacements: ReadonlyMap<string, Uint8Array>,
  applied: Set<string>,
): CfbNode[] {
  return nodes.map((node) => {
    if ('data' in node) {
      const data = replacements.get(node.name);
      if (data !== undefined) {
        applied.add(node.name);
        return {name: node.name, data};
      }
      return node;
    }
    return {name: node.name, children: replaceStreams(node.children, replacements, applied)};
  });
}
