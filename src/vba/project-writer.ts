// Synthesize a `vbaProject.bin` from module source — §2.3c of the VBA authoring line.
//
// Assembles the streams a macro-enabled workbook needs from the two encode primitives (writeCompoundFile
// + compressContainer): the `dir` record stream ([MS-OVBA] 2.3.4.2), the `_VBA_PROJECT` version header,
// the `PROJECT`/`PROJECTwm` text streams, and one compressed source stream per module, all nested under a
// `VBA` storage. Each module carries its source at MODULEOFFSET 0 with no p-code; the `_VBA_PROJECT`
// header advertises a version the host will not match, so Excel discards the (absent) PerformanceCache
// and recompiles from source on open — the sanctioned way to author without emitting a p-code cache.
//
// The output is validated against real Excel (open-clean + macro-enabled re-save preserves every module;
// see ADR 0017 §2.3c), not merely against this library's own reader.

import {type CfbNode, writeCompoundFile} from './cfb-writer.ts';
import {type Encoder, encoderForCodePage} from './codepage.ts';
import {VbaAuthorError} from './errors.ts';
import {compressContainer} from './ms-ovba.ts';

/** A module to author. `document` and `designer` kinds are host-coupled and not yet synthesizable. */
export interface VbaModuleSource {
  /** The module's code name — a valid VBA identifier, at most 31 characters. */
  readonly name: string;
  /** Procedural (`.bas`, a standard module) or a class module. */
  readonly kind: 'procedural' | 'class';
  /** The VBA source text. Newlines should be CRLF, as the VBA editor uses. */
  readonly source: string;
}

export interface VbaProjectSpec {
  readonly modules: readonly VbaModuleSource[];
  /** Project code page (`PROJECTCODEPAGE`). Defaults to 1252 (Western European). */
  readonly codePage?: number;
  /** The `PROJECTNAME` shown in the VBA editor. A valid identifier; defaults to `VBAProject`. */
  readonly projectName?: string;
}

const DEFAULT_CODE_PAGE = 1252;
const DEFAULT_PROJECT_NAME = 'VBAProject';
const IDENTIFIER = /^[A-Za-z][A-Za-z0-9_]*$/;
const MAX_NAME_CHARS = 31; // the CFB stream-name limit, and VBA's own module-name limit

// `dir`-stream record ids we emit ([MS-OVBA] 2.3.4.2), mirroring what Excel writes for a reference-free
// project. Ids are shared with the reader (project.ts) but kept local: the writer emits a superset.
const REC = {
  SYSKIND: 0x0001,
  LCID: 0x0002,
  LCIDINVOKE: 0x0014,
  CODEPAGE: 0x0003,
  NAME: 0x0004,
  DOCSTRING: 0x0005,
  DOCSTRING_UNICODE: 0x0040,
  HELPFILE1: 0x0006,
  HELPFILE2: 0x003d,
  HELPCONTEXT: 0x0007,
  LIBFLAGS: 0x0008,
  VERSION: 0x0009,
  CONSTANTS: 0x000c,
  CONSTANTS_UNICODE: 0x003c,
  MODULES_COUNT: 0x000f,
  PROJECTCOOKIE: 0x0013,
  MODULE_NAME: 0x0019,
  MODULE_NAME_UNICODE: 0x0047,
  MODULE_STREAMNAME: 0x001a,
  MODULE_STREAMNAME_UNICODE: 0x0032,
  MODULE_DOCSTRING: 0x001c,
  MODULE_DOCSTRING_UNICODE: 0x0048,
  MODULE_OFFSET: 0x0031,
  MODULE_HELPCONTEXT: 0x001e,
  MODULE_COOKIE: 0x002c,
  MODULE_TYPE_PROCEDURAL: 0x0021,
  MODULE_TYPE_OTHER: 0x0022,
  MODULE_TERMINATOR: 0x002b,
  TERMINATOR: 0x0010,
} as const;

const SYSKIND_WIN32 = 0x01;
const LCID_EN_US = 0x0409;
const COOKIE_NONE = 0xffff;
// _VBA_PROJECT header: Reserved1 (0x61CC) · Version (0xFFFF, unmatchable → forces recompile) · Reserved2
// (0x00) · Reserved3 (0x0000), with no trailing PerformanceCache. ([MS-OVBA] 2.3.4.1)
const VBA_PROJECT_HEADER = Uint8Array.from([0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00]);

/**
 * Synthesize a `vbaProject.bin` from `spec`. The returned bytes can be attached to a workbook via
 * `Workbook.vbaProjectBytes` (or re-parsed with `parseVbaProject`); the written workbook opens in Excel
 * as a macro-enabled book whose modules recompile from the embedded source.
 *
 * @throws {VbaAuthorError} on an invalid or duplicate module name, an unsupported module kind, or source
 *   / a name that the chosen code page cannot represent.
 */
export function writeVbaProject(spec: VbaProjectSpec): Uint8Array {
  const codePage = spec.codePage ?? DEFAULT_CODE_PAGE;
  const projectName = spec.projectName ?? DEFAULT_PROJECT_NAME;
  const encode = encoderForCodePage(codePage);

  validate(spec.modules, projectName);

  const vbaChildren: CfbNode[] = [
    {name: 'dir', data: compressContainer(buildDir(spec.modules, codePage, projectName, encode))},
    {name: '_VBA_PROJECT', data: VBA_PROJECT_HEADER},
  ];
  for (const m of spec.modules) {
    vbaChildren.push({name: m.name, data: compressContainer(encode(m.source))});
  }

  return writeCompoundFile([
    {name: 'PROJECT', data: buildProjectStream(spec.modules, projectName, encode)},
    {name: 'PROJECTwm', data: buildProjectwm(spec.modules, encode)},
    {name: 'VBA', children: vbaChildren},
  ]);
}

function validate(modules: readonly VbaModuleSource[], projectName: string): void {
  if (!IDENTIFIER.test(projectName) || projectName.length > MAX_NAME_CHARS) {
    throw new VbaAuthorError(
      `invalid project name '${projectName}' (must be a VBA identifier ≤ 31 chars)`,
    );
  }
  const seen = new Set<string>();
  for (const m of modules) {
    if (!IDENTIFIER.test(m.name) || m.name.length > MAX_NAME_CHARS) {
      throw new VbaAuthorError(
        `invalid module name '${m.name}' (must be a VBA identifier ≤ 31 chars)`,
      );
    }
    const key = m.name.toUpperCase(); // VBA names are case-insensitive
    if (seen.has(key)) throw new VbaAuthorError(`duplicate module name '${m.name}'`);
    seen.add(key);
    if (m.kind !== 'procedural' && m.kind !== 'class') {
      throw new VbaAuthorError(`module '${m.name}': kind '${m.kind}' is not yet synthesizable`);
    }
  }
}

function buildDir(
  modules: readonly VbaModuleSource[],
  codePage: number,
  projectName: string,
  encode: Encoder,
): Uint8Array {
  const r: number[] = [];
  push(r, REC.SYSKIND, u32(SYSKIND_WIN32));
  push(r, REC.LCID, u32(LCID_EN_US));
  push(r, REC.LCIDINVOKE, u32(LCID_EN_US));
  push(r, REC.CODEPAGE, u16(codePage));
  push(r, REC.NAME, [...encode(projectName)]);
  push(r, REC.DOCSTRING, []);
  push(r, REC.DOCSTRING_UNICODE, []);
  push(r, REC.HELPFILE1, []);
  push(r, REC.HELPFILE2, []);
  push(r, REC.HELPCONTEXT, u32(0));
  push(r, REC.LIBFLAGS, u32(0));
  // PROJECTVERSION: the Size field counts only the 4-byte VersionMajor; the 2-byte VersionMinor trails
  // it uncounted (the record that misaligns a naive TLV walk — see the reader).
  push(r, REC.VERSION, u32(0));
  r.push(...u16(3));
  push(r, REC.CONSTANTS, []);
  push(r, REC.CONSTANTS_UNICODE, []);
  // No PROJECTREFERENCES: a reference-free project opens clean in Excel, which re-adds host defaults.
  push(r, REC.MODULES_COUNT, u16(modules.length));
  push(r, REC.PROJECTCOOKIE, u16(COOKIE_NONE));
  for (const m of modules) {
    const nameBytes = [...encode(m.name)];
    const nameUnicode = utf16le(m.name);
    push(r, REC.MODULE_NAME, nameBytes);
    push(r, REC.MODULE_NAME_UNICODE, nameUnicode);
    push(r, REC.MODULE_STREAMNAME, nameBytes);
    push(r, REC.MODULE_STREAMNAME_UNICODE, nameUnicode);
    push(r, REC.MODULE_DOCSTRING, []);
    push(r, REC.MODULE_DOCSTRING_UNICODE, []);
    push(r, REC.MODULE_OFFSET, u32(0)); // source at stream start, no p-code prefix → recompile
    push(r, REC.MODULE_HELPCONTEXT, u32(0));
    push(r, REC.MODULE_COOKIE, u16(COOKIE_NONE));
    push(r, m.kind === 'procedural' ? REC.MODULE_TYPE_PROCEDURAL : REC.MODULE_TYPE_OTHER, []);
    push(r, REC.MODULE_TERMINATOR, []);
  }
  r.push(...u16(REC.TERMINATOR), ...u32(0)); // dir Terminator + Reserved
  return Uint8Array.from(r);
}

function buildProjectStream(
  modules: readonly VbaModuleSource[],
  projectName: string,
  encode: Encoder,
): Uint8Array {
  const lines = ['ID="{00000000-0000-0000-0000-000000000000}"'];
  for (const m of modules) {
    lines.push(`${m.kind === 'procedural' ? 'Module' : 'Class'}=${m.name}`);
  }
  lines.push(`Name="${projectName}"`, 'HelpContextID="0"', 'VersionCompatible32="393222000"');
  // CMG/DPB/GC hold the project's (encrypted) protection state; Excel regenerates them on save, and an
  // unprotected reference-free project opens clean without them.
  lines.push('CMG=""', 'DPB=""', 'GC=""', '');
  lines.push(
    '[Host Extender Info]',
    '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
    '',
    '[Workspace]',
  );
  for (const m of modules) lines.push(`${m.name}=0, 0, 0, 0, C`);
  lines.push('');
  return encode(lines.join('\r\n'));
}

// The PROJECTwm stream pairs each module's MBCS name with its UTF-16 name, both NUL-terminated, and ends
// with an empty pair.
function buildProjectwm(modules: readonly VbaModuleSource[], encode: Encoder): Uint8Array {
  const b: number[] = [];
  for (const m of modules) {
    b.push(...encode(m.name), 0x00, ...utf16le(m.name), 0x00, 0x00);
  }
  b.push(0x00, 0x00);
  return Uint8Array.from(b);
}

function push(out: number[], id: number, data: number[]): void {
  out.push(...u16(id), ...u32(data.length), ...data);
}
function u16(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff];
}
function u32(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
}
function utf16le(s: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < s.length; i++) out.push(...u16(s.charCodeAt(i)));
  return out;
}
