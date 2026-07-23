// Turn a raw `vbaProject.bin` into readable module source.
//
// Pipeline ([MS-OVBA] 2.3.4.2): the CFB container holds a `VBA/dir` stream (itself MS-OVBA compressed)
// that lists each module's stream name, its code page, and the byte offset in that module's stream
// where the compressed source text begins — past the p-code / PerformanceCache. We decompress `dir`,
// read those facts, then decompress each module stream from its text offset and decode with the
// project code page. The p-code is version-specific and deliberately not exposed; a reader wants source.

import {CompoundFile} from './cfb.ts';
import {type Decoder, decoderForCodePage} from './codepage.ts';
import {VbaParseError} from './errors.ts';
import {decompressContainer} from './ms-ovba.ts';

/** How a module participates in the project — the classification the VBA editor shows. */
export type VbaModuleKind = 'procedural' | 'document' | 'class' | 'designer';

export interface VbaModule {
  /** The module's code name as seen in the VBA editor, e.g. `ThisWorkbook`, `JsonConverter`. */
  readonly name: string;
  /** The CFB stream the module's bytes live in — usually equal to {@link name}. */
  readonly streamName: string;
  /** Procedural (`.bas`), document code-behind, class module, or designer (UserForm). */
  readonly kind: VbaModuleKind;
  /** The decompressed VBA source (p-code and PerformanceCache are not included). */
  readonly source: string;
}

export interface VbaProject {
  /** The project code page (`PROJECTCODEPAGE`) used to decode module names and source. */
  readonly codePage: number;
  /** The project's modules, in declaration order. */
  readonly modules: readonly VbaModule[];
}

// `dir`-stream record ids we consume ([MS-OVBA] 2.3.4.2). Every other record is skipped by the uniform
// TLV walk; its Size field already accounts for its payload, so skipping is just advancing past it.
const REC_PROJECT_CODEPAGE = 0x0003;
const REC_MODULE_NAME = 0x0019;
const REC_MODULE_STREAMNAME = 0x001a;
const REC_MODULE_TYPE_PROCEDURAL = 0x0021;
const REC_MODULE_TYPE_DOCUMENT = 0x0022;
const REC_MODULE_OFFSET = 0x0031;
const REC_MODULE_TERMINATOR = 0x002b;
// PROJECTVERSION carries VersionMajor (u32) + VersionMinor (u16) after its fixed Size=4 field, but Size
// only accounts for the 4-byte major. The extra 2-byte minor is uncounted, so a uniform TLV walk
// misaligns here by 2 bytes unless it is skipped explicitly.
const REC_PROJECT_VERSION = 0x0009;

interface PendingModule {
  name?: Uint8Array;
  streamName?: Uint8Array;
  offset?: number;
  documentType?: boolean;
}

export function parseVbaProject(bin: Uint8Array): VbaProject {
  const cfb = new CompoundFile(bin);

  const dirCompressed = cfb.readStream('dir');
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const dir = decompressContainer(dirCompressed);

  let codePage = 1252; // Western-European default until PROJECTCODEPAGE says otherwise.
  const rawModules: PendingModule[] = [];
  let pending: PendingModule = {};

  // The dir stream is a flat sequence of TLV records: Id(u16) Size(u32) data[Size].
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const id = readU16(dir, pos);
    const size = readU32(dir, pos + 2);
    const dataStart = pos + 6;
    if (dataStart + size > dir.length)
      throw new VbaParseError(`dir record 0x${id.toString(16)} overruns stream`);
    pos = dataStart + size;
    if (id === REC_PROJECT_VERSION) pos += 2; // uncounted VersionMinor (u16)

    switch (id) {
      case REC_PROJECT_CODEPAGE:
        if (size >= 2) codePage = readU16(dir, dataStart);
        break;
      case REC_MODULE_NAME:
        pending.name = dir.subarray(dataStart, dataStart + size);
        break;
      case REC_MODULE_STREAMNAME:
        pending.streamName = dir.subarray(dataStart, dataStart + size);
        break;
      case REC_MODULE_TYPE_PROCEDURAL:
        pending.documentType = false;
        break;
      case REC_MODULE_TYPE_DOCUMENT:
        pending.documentType = true;
        break;
      case REC_MODULE_OFFSET:
        if (size >= 4) pending.offset = readU32(dir, dataStart);
        break;
      case REC_MODULE_TERMINATOR:
        if (pending.streamName !== undefined && pending.offset !== undefined)
          rawModules.push(pending);
        pending = {};
        break;
      default:
        break;
    }
  }

  const decoder = decoderForCodePage(codePage);
  const kindByName = readProjectStreamKinds(cfb, decoder);

  const modules: VbaModule[] = rawModules.map((m) => {
    const streamName = decoder.decode(m.streamName);
    const name = m.name !== undefined ? decoder.decode(m.name) : streamName;
    const kind = kindByName.get(name) ?? (m.documentType ? 'document' : 'procedural');
    return {
      name,
      streamName,
      kind,
      source: readModuleSource(cfb, streamName, m.offset as number, decoder),
    };
  });

  return {codePage, modules};
}

function readModuleSource(
  cfb: CompoundFile,
  streamName: string,
  textOffset: number,
  decoder: Decoder,
): string {
  const stream = cfb.readStream(streamName);
  if (!stream) throw new VbaParseError(`module stream '${streamName}' not found in container`);
  return decoder.decode(decompressContainer(stream, textOffset));
}

// The `PROJECT` stream (uncounted, plain MBCS text) declares each module's exact kind by keyword:
// `Document=`, `Module=` (procedural .bas), `Class=`, `BaseClass=` (designer/UserForm). The `dir`
// stream's MODULETYPE only distinguishes procedural from non-procedural, so PROJECT refines it. Absent
// or unparsable PROJECT just falls back to that coarser MODULETYPE classification.
function readProjectStreamKinds(cfb: CompoundFile, decoder: Decoder): Map<string, VbaModuleKind> {
  const kinds = new Map<string, VbaModuleKind>();
  const stream = cfb.readStream('PROJECT');
  if (!stream) return kinds;
  const text = decoder.decode(stream);
  const keyword: Record<string, VbaModuleKind> = {
    Document: 'document',
    Module: 'procedural',
    Class: 'class',
    BaseClass: 'designer',
  };
  for (const line of text.split(/\r\n|\r|\n/)) {
    const eq = line.indexOf('=');
    if (eq < 0) continue;
    const kind = keyword[line.slice(0, eq)];
    if (kind === undefined) continue;
    // Value is `Name` or `Name/&H00000000` (document modules carry a cookie); take the name.
    const name = (line.slice(eq + 1).split('/')[0] ?? '').trim();
    if (name) kinds.set(name, kind);
  }
  return kinds;
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
