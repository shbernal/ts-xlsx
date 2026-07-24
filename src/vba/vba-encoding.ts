// Shared low-level encoders for the VBA `dir` stream and name validation — the primitives the structural
// splices in `project-editor.ts` build on. ([MS-OVBA] 2.3.4.2 record TLVs, and the VBA identifier rules.)
//
// There is no from-scratch `vbaProject.bin` synthesizer here. Excel does not recompile VBA from source
// on open — a module runs the compiled p-code it ships — so authoring/editing module SOURCE is done by
// the offline `tools/vba-compiler` (VBIDE), which produces genuinely compiled p-code. This module holds
// only what the pure-TS structural edits (remove module, add reference) still need (ADR 0019).

import {VbaAuthorError} from './errors.ts';

const IDENTIFIER = /^[A-Za-z][A-Za-z0-9_]*$/;
const MAX_NAME_CHARS = 31; // the CFB stream-name limit, and VBA's own module-name limit

/**
 * Validate a module, project, or reference name against the shared VBA identifier contract: a valid
 * identifier, at most 31 characters (the CFB stream-name limit, which doubles as VBA's own module-name
 * limit). Used by {@link project-editor.ts | project-editor}'s structural edits.
 *
 * @throws {VbaAuthorError} if `name` is not a valid VBA identifier or exceeds 31 characters.
 */
export function validateVbaName(name: string, what: 'project' | 'module' | 'reference'): void {
  if (!IDENTIFIER.test(name) || name.length > MAX_NAME_CHARS) {
    throw new VbaAuthorError(
      `invalid ${what} name '${name}' (must be a VBA identifier ≤ 31 chars)`,
    );
  }
}

/** Append one `dir`-stream TLV record (Id, Size, data) to `out`. */
export function push(out: number[], id: number, data: number[]): void {
  out.push(...u16(id), ...u32(data.length), ...data);
}
export function u16(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff];
}
export function u32(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
}
/** A name as NUL-free UTF-16LE code units — the encoding [MS-OVBA] uses for every "Unicode" name field. */
export function utf16le(s: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < s.length; i++) out.push(...u16(s.charCodeAt(i)));
  return out;
}
