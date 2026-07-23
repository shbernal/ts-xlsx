// Edit an existing `vbaProject.bin` in place — swap one or more modules' source, preserve everything
// else. The complement to writeVbaProject (§2.3c), which synthesizes a *reference-free, host-default*
// project from scratch: a from-scratch author drops the PROJECTREFERENCES, host-extender info, and
// project constants a real project carries. Editing must keep them.
//
// So this is a surgical splice, not a re-synthesis. We parse the original container, rebuild its whole
// storage/stream tree, and replace exactly three things: (1) each edited module's stream — the freshly
// MS-OVBA-compressed new source, at MODULEOFFSET 0 with no p-code; (2) that module's MODULEOFFSET record
// in the `dir` stream, set to 0; (3) the `_VBA_PROJECT` header, set to the unmatchable-version cookie so
// Excel discards every module's stale PerformanceCache and recompiles all of them from source. The
// `dir` reference records, the `PROJECT`/`PROJECTwm` text, designer storages, and every untouched
// module's bytes ride through the tree unchanged — preservation is by *not touching* them.
//
// Because host linkage is inherited from the preserved streams rather than synthesized, this can edit a
// `document` or `designer` module (ThisWorkbook, Sheet1, a UserForm) that writeVbaProject rejects.

import {CompoundFile} from './cfb.ts';
import type {CfbNode, CfbStream} from './cfb-writer.ts';
import {writeCompoundFile} from './cfb-writer.ts';
import {decoderForCodePage, type Encoder, encoderForCodePage} from './codepage.ts';
import {VbaAuthorError, VbaParseError} from './errors.ts';
import {compressContainer, decompressContainer} from './ms-ovba.ts';
import {parseVbaProject} from './project.ts';
import {
  buildModuleDirRecord,
  push,
  u16,
  u32,
  utf16le,
  type VbaModuleSource,
  validateVbaName,
} from './project-writer.ts';

const DIR_STREAM = 'dir';
const VBA_PROJECT_STREAM = '_VBA_PROJECT';
const PROJECT_STREAM = 'PROJECT';
const PROJECTWM_STREAM = 'PROJECTwm';
const VBA_STORAGE = 'VBA';

// `dir`-record ids this splice reads to locate a module's offset field ([MS-OVBA] 2.3.4.2); every other
// record is preserved verbatim.
const REC_MODULE_STREAMNAME = 0x001a;
const REC_MODULE_OFFSET = 0x0031;
const REC_MODULE_TERMINATOR = 0x002b;
const REC_PROJECT_VERSION = 0x0009; // its uncounted 2-byte VersionMinor trails the counted payload
// `dir`-record ids the add-module splice also reads: MODULES_COUNT to bump the module tally, and the
// PROJECTMODULES terminator (distinct from a per-module MODULETERMINATOR) that closes the whole `dir`
// stream — a new module's record block is inserted right before it.
const REC_MODULES_COUNT = 0x000f;
const REC_DIR_TERMINATOR = 0x0010;

// `dir`-record ids the add-reference splice builds ([MS-OVBA] 2.3.4.2.2). REFERENCENAME's Unicode half
// is a *literal* 0x003E marker, not a nested record id — but it is laid out as its own Id+Size+data TLV,
// so a generic walk (and this splice) sees REFERENCENAME as two chained records, exactly like
// MODULE_NAME/MODULE_NAME_UNICODE. Verified against a real Excel-authored dir stream (2026-07-23).
const REC_REFERENCE_NAME = 0x0016;
const REC_REFERENCE_NAME_UNICODE = 0x003e;
const REC_REFERENCE_REGISTERED = 0x000d;

// _VBA_PROJECT header advertising a version the host cannot match, so Excel drops all PerformanceCache
// and recompiles every module from source — the same cookie writeVbaProject emits ([MS-OVBA] 2.3.4.1).
const RECOMPILE_HEADER = Uint8Array.from([0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00]);

/**
 * Edit the source of one or more existing modules in a `vbaProject.bin`, returning new bytes that
 * preserve the project's references, host info, and every other module. `edits` maps a module's code
 * name (case-insensitively, as VBA compares them) to its replacement source.
 *
 * @throws {VbaParseError} if `bin` is not a parseable VBA project (validated before any edit).
 * @throws {VbaAuthorError} if a named module is absent, or the new source has a character the project's
 *   code page cannot represent.
 */
export function editVbaModuleSources(
  bin: Uint8Array,
  edits: ReadonlyMap<string, string>,
): Uint8Array {
  if (edits.size === 0) return bin;

  // Parse fail-closed first: this validates the container and resolves each module name to its stream
  // name and the project code page, so nothing is mutated on a bad input or an unknown module.
  const project = parseVbaProject(bin);
  const moduleByName = new Map(project.modules.map((m) => [m.name.toUpperCase(), m]));
  const newSourceByStream = new Map<string, string>();
  for (const [name, source] of edits) {
    const module = moduleByName.get(name.toUpperCase());
    if (!module) throw new VbaAuthorError(`module '${name}' is not in the VBA project`);
    if (module.streamName === DIR_STREAM || module.streamName === VBA_PROJECT_STREAM) {
      throw new VbaAuthorError(
        `cannot edit module with reserved stream name '${module.streamName}'`,
      );
    }
    newSourceByStream.set(module.streamName, source);
  }

  const encode = encoderForCodePage(project.codePage);
  const cfb = new CompoundFile(bin);

  const dirCompressed = cfb.readStream(DIR_STREAM);
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const patchedDir = zeroModuleOffsets(
    decompressContainer(dirCompressed),
    new Set(newSourceByStream.keys()),
    project.codePage,
  );

  const replacements = new Map<string, Uint8Array>([
    [DIR_STREAM, compressContainer(patchedDir)],
    [VBA_PROJECT_STREAM, RECOMPILE_HEADER],
  ]);
  for (const [streamName, source] of newSourceByStream) {
    replacements.set(streamName, compressContainer(encode(source)));
  }

  const applied = new Set<string>();
  const newTree = replaceStreams(cfb.tree(), replacements, applied);

  // `dir` and every edited module must have been present in the tree; `_VBA_PROJECT` is expected but a
  // minimal/synthetic project may omit it, so its absence is tolerated (recompile is moot without it).
  if (!applied.has(DIR_STREAM))
    throw new VbaParseError("VBA project 'dir' stream is not in the container tree");
  for (const streamName of newSourceByStream.keys()) {
    if (!applied.has(streamName)) {
      throw new VbaAuthorError(`module stream '${streamName}' is not in the container tree`);
    }
  }

  return writeCompoundFile(newTree);
}

/**
 * Add a standard module to an existing `vbaProject.bin`, returning new bytes that carry every existing
 * module, reference, and host-info record unchanged. This is the splice counterpart to
 * {@link editVbaModuleSources}: rather than swapping a module's source, it grows the project by one
 * module — a new `VBA/<name>` stream, a new MODULE record block in `dir` (with `MODULES_COUNT`
 * incremented), and the matching `Module=`/`Class=` and workspace lines in `PROJECT`/`PROJECTwm`.
 *
 * Only `procedural` and `class` modules can be added this way — a `document` or `designer` module's
 * name must correspond to a worksheet/workbook `codeName` the host already knows about, which this
 * project-level primitive has no visibility into (see `writeVbaProject`, which rejects the same kinds
 * for the same reason).
 *
 * @throws {VbaParseError} if `bin` is not a parseable VBA project (validated before any edit).
 * @throws {VbaAuthorError} if `module.name` is not a valid VBA identifier, collides (case-insensitively)
 *   with an existing module or a reserved stream name, or `module.source` has a character the project's
 *   code page cannot represent.
 */
export function addVbaModule(bin: Uint8Array, module: VbaModuleSource): Uint8Array {
  validateVbaName(module.name, 'module');
  if (module.kind !== 'procedural' && module.kind !== 'class') {
    throw new VbaAuthorError(`module '${module.name}': kind '${module.kind}' is not yet addable`);
  }

  // Parse fail-closed first: validates the container and resolves the existing module names, so nothing
  // is mutated on a bad input or a name collision.
  const project = parseVbaProject(bin);
  const nameKey = module.name.toUpperCase(); // VBA names are case-insensitive
  if (nameKey === DIR_STREAM.toUpperCase() || nameKey === VBA_PROJECT_STREAM.toUpperCase()) {
    throw new VbaAuthorError(`module name '${module.name}' collides with a reserved stream name`);
  }
  if (project.modules.some((m) => m.name.toUpperCase() === nameKey)) {
    throw new VbaAuthorError(`module '${module.name}' already exists in the project`);
  }

  const encode = encoderForCodePage(project.codePage);
  const cfb = new CompoundFile(bin);

  const dirCompressed = cfb.readStream(DIR_STREAM);
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const patchedDir = insertModuleDirRecord(decompressContainer(dirCompressed), module, encode);

  const replacements = new Map<string, Uint8Array>([[DIR_STREAM, compressContainer(patchedDir)]]);
  if (cfb.readStream(VBA_PROJECT_STREAM)) replacements.set(VBA_PROJECT_STREAM, RECOMPILE_HEADER);

  const projectText = cfb.readStream(PROJECT_STREAM);
  if (projectText) {
    const decoder = decoderForCodePage(project.codePage);
    replacements.set(
      PROJECT_STREAM,
      encode(insertProjectStreamLines(decoder.decode(projectText), module)),
    );
  }
  const projectwm = cfb.readStream(PROJECTWM_STREAM);
  if (projectwm) {
    replacements.set(
      PROJECTWM_STREAM,
      insertProjectwmRecord(projectwm, project.modules.length, module.name, encode),
    );
  }

  const applied = new Set<string>();
  const withReplacements = replaceStreams(cfb.tree(), replacements, applied);
  if (!applied.has(DIR_STREAM))
    throw new VbaParseError("VBA project 'dir' stream is not in the container tree");

  const moduleNode: CfbStream = {
    name: module.name,
    data: compressContainer(encode(module.source)),
  };
  const inserted = new Set<string>();
  const newTree = insertIntoStorage(withReplacements, VBA_STORAGE, moduleNode, inserted);
  if (!inserted.has(VBA_STORAGE)) {
    throw new VbaParseError(`VBA project has no '${VBA_STORAGE}' storage to add the module to`);
  }

  return writeCompoundFile(newTree);
}

/**
 * A registered (COM Automation type-library) reference to add to an existing VBA project — the shape of
 * a real "add a reference to Microsoft Scripting Runtime" call. Project references (to another VBA
 * project) and control references (to an ActiveX control library) are out of scope — see
 * {@link addVbaReference}.
 */
export interface VbaLibraryReference {
  /**
   * The reference's namespace name in the VBA editor — what a qualified reference like
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
  /** The type library's major version — an integer in `[0, 0xFFFF]` ([MS-OVBA] `LibidMajorVersion`). */
  readonly majorVersion: number;
  /** The type library's minor version — an integer in `[0, 0xFFFF]` ([MS-OVBA] `LibidMinorVersion`). */
  readonly minorVersion: number;
  /**
   * The type library's LCID — an integer in `[0, 0xFFFFFFFF]`. Defaults to `0` (locale-neutral), the
   * overwhelming common case (every reference in a real project observed while building this had `0`).
   */
  readonly lcid?: number;
  /** Absolute Windows path to the type library file, e.g. `C:\Windows\System32\scrrun.dll`. */
  readonly path: string;
}

// LibidMajorVersion/LibidMinorVersion ([MS-OVBA] 2.1.1.8): 1*4HEXDIG, so at most 0xFFFF.
const MAX_LIBID_VERSION = 0xffff;
// LibidLcid: 1*8HEXDIG, so at most 0xFFFFFFFF (practically always 0 — locale-neutral).
const MAX_LIBID_LCID = 0xffffffff;
// LibidRegName: *255(%x01-FF) — at most 255 bytes, never NUL.
const MAX_DISPLAY_NAME_CHARS = 255;
const GUID_PATTERN =
  /^\{?([0-9A-Fa-f]{8})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{12})\}?$/;

interface NormalizedReference {
  readonly name: string;
  readonly libid: string;
}

// Validate every field fail-closed and assemble the Libid string ([MS-OVBA] 2.1.1.8 LibidReference ABNF:
// `*\G{GUID}#Major.Minor#LCID#Path#RegName`, hex digit strings with no `0x` prefix) — confirmed
// byte-for-byte against a real Excel-authored reference (2026-07-23):
// `*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\System32\scrrun.dll#Microsoft Scripting Runtime`.
function normalizeReference(ref: VbaLibraryReference): NormalizedReference {
  validateVbaName(ref.name, 'reference');

  const guidMatch = GUID_PATTERN.exec(ref.guid.trim());
  if (!guidMatch) throw new VbaAuthorError(`invalid reference GUID '${ref.guid}'`);
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
      `invalid reference path '${ref.path}' (must be non-empty and contain no NUL or '#')`,
    );
  }
  const displayName = ref.displayName ?? ref.name;
  if (
    displayName.length === 0 ||
    displayName.length > MAX_DISPLAY_NAME_CHARS ||
    displayName.includes('\0')
  ) {
    throw new VbaAuthorError(`invalid reference display name '${displayName}'`);
  }

  const libid =
    `*\\G${guid}#${ref.majorVersion.toString(16).toUpperCase()}.` +
    `${ref.minorVersion.toString(16).toUpperCase()}#${lcid.toString(16).toUpperCase()}` +
    `#${ref.path}#${displayName}`;
  return {name: ref.name, libid};
}

/**
 * Add a registered (COM type-library) reference to an existing `vbaProject.bin`, returning new bytes
 * that carry every existing module, reference, and host-info record unchanged. The splice counterpart to
 * {@link addVbaModule}: it grows the project's `dir` stream by one `REFERENCENAME` + `REFERENCEREGISTERED`
 * record pair, positioned immediately before `MODULES_COUNT` (references have no count field of their
 * own — `MODULES_COUNT` simply marks where the reference array ends). Unlike adding a module, this needs
 * no change to `PROJECT`/`PROJECTwm`: a real Excel-authored `PROJECT` stream carries no `Reference=` line
 * at all — references live only in `dir` (confirmed against a genuine Excel-authored project).
 *
 * @throws {VbaParseError} if `bin` is not a parseable VBA project (validated before any edit).
 * @throws {VbaAuthorError} if any field of `ref` is invalid (see {@link VbaLibraryReference}), or the
 *   assembled reference text has a character the project's code page cannot represent.
 */
export function addVbaReference(bin: Uint8Array, ref: VbaLibraryReference): Uint8Array {
  const normalized = normalizeReference(ref);

  // Parse fail-closed first: validates the container before any mutation.
  const project = parseVbaProject(bin);
  const encode = encoderForCodePage(project.codePage);
  const cfb = new CompoundFile(bin);

  const dirCompressed = cfb.readStream(DIR_STREAM);
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const records = buildReferenceDirRecords(normalized, encode);
  const patchedDir = insertReferenceDirRecords(decompressContainer(dirCompressed), records);

  const replacements = new Map<string, Uint8Array>([[DIR_STREAM, compressContainer(patchedDir)]]);
  if (cfb.readStream(VBA_PROJECT_STREAM)) replacements.set(VBA_PROJECT_STREAM, RECOMPILE_HEADER);

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

// Insert new reference dir records right before MODULES_COUNT (0x000f) — the reference array has no
// explicit count field; MODULES_COUNT is simply the next record once the last reference ends (confirmed
// against a real Excel-authored dir stream). Every other record — other references, all modules — rides
// through unchanged.
function insertReferenceDirRecords(dir: Uint8Array, records: readonly number[]): Uint8Array {
  let insertAt = -1;
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const recordStart = pos;
    const id = readU16(dir, pos);
    const size = readU32(dir, pos + 2);
    const dataStart = pos + 6;
    if (dataStart + size > dir.length) {
      throw new VbaParseError(`dir record 0x${id.toString(16)} overruns while adding a reference`);
    }
    pos = dataStart + size;
    if (id === REC_PROJECT_VERSION) pos += 2; // uncounted VersionMinor (u16)
    if (id === REC_MODULES_COUNT) {
      insertAt = recordStart;
      break;
    }
  }
  if (insertAt < 0) throw new VbaParseError('dir stream is missing MODULES_COUNT');

  const rec = Uint8Array.from(records);
  const out = new Uint8Array(dir.length + rec.length);
  out.set(dir.subarray(0, insertAt), 0);
  out.set(rec, insertAt);
  out.set(dir.subarray(insertAt), insertAt + rec.length);
  return out;
}

// Insert `node` as a new child of the first storage named `storageName` found in the tree (depth-first),
// marking `storageName` in `applied` once done. Used to add a module's stream alongside `dir` and
// `_VBA_PROJECT`, which already live in the `VBA` storage.
function insertIntoStorage(
  nodes: readonly CfbNode[],
  storageName: string,
  node: CfbNode,
  applied: Set<string>,
): CfbNode[] {
  return nodes.map((n) => {
    if ('data' in n) return n;
    const children = insertIntoStorage(n.children, storageName, node, applied);
    if (n.name === storageName && !applied.has(storageName)) {
      applied.add(storageName);
      return {name: n.name, children: [...children, node]};
    }
    return {name: n.name, children};
  });
}

// Insert a new MODULE record block into a decompressed `dir` stream, right before the PROJECTMODULES
// terminator (0x0010) that closes the structure, and increment MODULES_COUNT. Every other record —
// PROJECTREFERENCES, other modules, project-level fields — is carried through untouched.
function insertModuleDirRecord(
  dir: Uint8Array,
  module: VbaModuleSource,
  encode: Encoder,
): Uint8Array {
  let countAt = -1;
  let insertAt = -1;
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const recordStart = pos;
    const id = readU16(dir, pos);
    const size = readU32(dir, pos + 2);
    const dataStart = pos + 6;
    if (dataStart + size > dir.length) {
      throw new VbaParseError(`dir record 0x${id.toString(16)} overruns while adding a module`);
    }
    pos = dataStart + size;
    if (id === REC_PROJECT_VERSION) pos += 2; // uncounted VersionMinor (u16)

    if (id === REC_MODULES_COUNT) {
      if (size < 2) throw new VbaParseError('PROJECTMODULES MODULES_COUNT record is malformed');
      countAt = dataStart;
    } else if (id === REC_DIR_TERMINATOR) {
      insertAt = recordStart; // nothing meaningful follows the structure terminator
      break;
    }
  }
  if (countAt < 0 || insertAt < 0) {
    throw new VbaParseError('dir stream is missing MODULES_COUNT or the PROJECTMODULES terminator');
  }

  const out = dir.slice();
  const newCount = readU16(out, countAt) + 1;
  out[countAt] = newCount & 0xff;
  out[countAt + 1] = (newCount >> 8) & 0xff;

  const record = Uint8Array.from(buildModuleDirRecord(module, encode));
  const result = new Uint8Array(out.length + record.length);
  result.set(out.subarray(0, insertAt), 0);
  result.set(record, insertAt);
  result.set(out.subarray(insertAt), insertAt + record.length);
  return result;
}

// Insert the new module's declaration line (`Module=`/`Class=`) after the last existing
// Document=/Module=/Class=/BaseClass= line, and its workspace line (`Name=0, 0, 0, 0, C`) after the last
// existing line in the `[Workspace]` section, if present. Every other line — CMG/DPB/GC protection state,
// Host Extender Info, project name — is left exactly as it was.
function insertProjectStreamLines(text: string, module: VbaModuleSource): string {
  const eol = text.includes('\r\n') ? '\r\n' : '\n';
  const lines = text.split(/\r\n|\r|\n/);
  const declLine = `${module.kind === 'procedural' ? 'Module' : 'Class'}=${module.name}`;
  const isDeclLine = (l: string) => /^(Document|Module|Class|BaseClass)=/.test(l);

  let lastDecl = -1;
  for (let i = 0; i < lines.length; i++) if (isDeclLine(lines[i] as string)) lastDecl = i;
  if (lastDecl >= 0) {
    lines.splice(lastDecl + 1, 0, declLine);
  } else {
    const idLine = lines.findIndex((l) => l.startsWith('ID='));
    lines.splice(idLine + 1, 0, declLine);
  }

  const wsIndex = lines.findIndex((l) => l.trim() === '[Workspace]');
  if (wsIndex >= 0) {
    let lastWs = wsIndex;
    for (let i = wsIndex + 1; i < lines.length; i++) {
      const l = lines[i] as string;
      if (l.trim() === '' || l.startsWith('[')) break;
      lastWs = i;
    }
    lines.splice(lastWs + 1, 0, `${module.name}=0, 0, 0, 0, C`);
  }

  return lines.join(eol);
}

// Insert the new module's (MBCS name, UTF-16 name) pair into the binary PROJECTwm stream, right before
// its terminator. `existingModuleCount` (from the already fail-closed-parsed project) lets the walk
// consume exactly that many records before expecting the terminator, rather than guessing where a
// variable-length record ends.
function insertProjectwmRecord(
  wm: Uint8Array,
  existingModuleCount: number,
  name: string,
  encode: Encoder,
): Uint8Array {
  let pos = 0;
  for (let i = 0; i < existingModuleCount; i++) {
    const mbcsEnd = wm.indexOf(0x00, pos);
    if (mbcsEnd < 0)
      throw new VbaParseError('PROJECTwm record is missing its MBCS name terminator');
    pos = mbcsEnd + 1;
    let utf16End = pos;
    while (utf16End + 1 < wm.length && (wm[utf16End] !== 0 || wm[utf16End + 1] !== 0))
      utf16End += 2;
    if (utf16End + 1 >= wm.length) {
      throw new VbaParseError('PROJECTwm record is missing its Unicode name terminator');
    }
    pos = utf16End + 2;
  }
  if (pos + 2 > wm.length || wm[pos] !== 0 || wm[pos + 1] !== 0) {
    throw new VbaParseError('PROJECTwm terminator not found at the expected position');
  }

  const record = Uint8Array.from([...encode(name), 0x00, ...utf16le(name), 0x00, 0x00]);
  const out = new Uint8Array(wm.length + record.length);
  out.set(wm.subarray(0, pos), 0);
  out.set(record, pos);
  out.set(wm.subarray(pos), pos + record.length);
  return out;
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

// Walk the decompressed `dir` and set MODULEOFFSET to 0 for every module whose stream is being replaced —
// the edited stream now holds source at offset 0 with no p-code prefix. Every other record (references,
// host info, untouched modules) is left exactly as it was.
function zeroModuleOffsets(
  dir: Uint8Array,
  streamNames: ReadonlySet<string>,
  codePage: number,
): Uint8Array {
  const out = dir.slice();
  const decoder = decoderForCodePage(codePage);
  let currentStream: string | undefined;
  let pos = 0;
  while (pos + 6 <= out.length) {
    const id = readU16(out, pos);
    const size = readU32(out, pos + 2);
    const dataStart = pos + 6;
    if (dataStart + size > out.length) {
      throw new VbaParseError(`dir record 0x${id.toString(16)} overruns while editing`);
    }
    pos = dataStart + size;
    if (id === REC_PROJECT_VERSION) pos += 2; // uncounted VersionMinor (u16)

    if (id === REC_MODULE_STREAMNAME) {
      currentStream = decoder.decode(out.subarray(dataStart, dataStart + size));
    } else if (id === REC_MODULE_OFFSET) {
      if (size >= 4 && currentStream !== undefined && streamNames.has(currentStream)) {
        out[dataStart] = 0;
        out[dataStart + 1] = 0;
        out[dataStart + 2] = 0;
        out[dataStart + 3] = 0;
      }
    } else if (id === REC_MODULE_TERMINATOR) {
      currentStream = undefined;
    }
  }
  return out;
}

function readU16(buf: Uint8Array, at: number): number {
  return (buf[at] as number) | ((buf[at + 1] as number) << 8);
}
function readU32(buf: Uint8Array, at: number): number {
  return (
    ((buf[at] as number) |
      ((buf[at + 1] as number) << 8) |
      ((buf[at + 2] as number) << 16) |
      ((buf[at + 3] as number) << 24)) >>>
    0
  );
}
