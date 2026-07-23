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
import type {CfbNode} from './cfb-writer.ts';
import {writeCompoundFile} from './cfb-writer.ts';
import {decoderForCodePage, encoderForCodePage} from './codepage.ts';
import {VbaAuthorError, VbaParseError} from './errors.ts';
import {compressContainer, decompressContainer} from './ms-ovba.ts';
import {parseVbaProject} from './project.ts';

const DIR_STREAM = 'dir';
const VBA_PROJECT_STREAM = '_VBA_PROJECT';

// `dir`-record ids this splice reads to locate a module's offset field ([MS-OVBA] 2.3.4.2); every other
// record is preserved verbatim.
const REC_MODULE_STREAMNAME = 0x001a;
const REC_MODULE_OFFSET = 0x0031;
const REC_MODULE_TERMINATOR = 0x002b;
const REC_PROJECT_VERSION = 0x0009; // its uncounted 2-byte VersionMinor trails the counted payload

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
