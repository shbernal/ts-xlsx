// Front-of-the-pipe format detection for the readers.
//
// The reader's first job on an arbitrary blob is to reject what it cannot read with a clear, typed
// error rather than a raw fflate crash. Two probes do it: a cheap magic-byte sniff *before* the zip
// layer runs (so a legacy `.xls` compound file or non-ZIP garbage fails fast, before fflate throws its
// opaque "end of central directory" error — which can also leak an absolute path from below), and a
// post-inflate check for the `.xlsb` binary workbook part (a real ZIP, but not XML). Both funnel into
// {@link UnsupportedFormatError} so callers branch on `.format`, never on message text.

import {UnsupportedFormatError} from './errors.ts';
import {inflatePackage} from './inflate.ts';

// The OLE2 / Compound File Binary signature ([MS-CFB] 2.2) that opens every legacy `.xls` (and the
// `.doc`/`.ppt` siblings). `src/vba/cfb.ts` reads this same magic for `vbaProject.bin`; here we need
// only recognise it, not parse the container.
const CFB_MAGIC = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);

// The ZIP local-file-header signature (`PK\x03\x04`) that opens every OPC package — `.xlsx`, `.xlsb`,
// and the rest. An empty or spanned archive starts with a different `PK` marker, but a real workbook
// package always leads with a local file header.
const ZIP_LOCAL_FILE_MAGIC = Uint8Array.of(0x50, 0x4b, 0x03, 0x04);

/** The container kind a leading magic-byte sniff distinguishes, before any inflation. */
export type Container = 'cfb' | 'zip' | 'other';

/** Classify a blob's container by its leading magic bytes alone — no allocation, no inflation. */
export function sniffContainer(data: Uint8Array): Container {
  if (startsWith(data, CFB_MAGIC)) return 'cfb';
  if (startsWith(data, ZIP_LOCAL_FILE_MAGIC)) return 'zip';
  return 'other';
}

function startsWith(data: Uint8Array, magic: Uint8Array): boolean {
  if (data.length < magic.length) return false;
  for (let i = 0; i < magic.length; i++) {
    if (data[i] !== magic[i]) return false;
  }
  return true;
}

/**
 * Inflate an `.xlsx` package, translating a non-`.xlsx` input into a typed {@link UnsupportedFormatError}
 * before or instead of a raw zip failure:
 * - a legacy `.xls` (CFB) blob → `'xls'`, caught by the magic sniff so fflate never runs on it;
 * - a non-ZIP blob → `'unknown'`, likewise caught before inflation;
 * - a `PK`-headed blob that fflate then rejects as malformed → `'unknown'`, with the underlying zip
 *   message discarded so no internals (or path) leak.
 *
 * The bounded-inflation guard (a probable zip bomb) is a legitimate, informative failure and is
 * re-thrown unchanged — it is not a format-classification error.
 */
export function inflateXlsxPackage(data: Uint8Array, cap: number): Record<string, Uint8Array> {
  const container = sniffContainer(data);
  if (container === 'cfb') throw new UnsupportedFormatError('xls');
  if (container === 'other') throw new UnsupportedFormatError('unknown');
  try {
    return inflatePackage(data, cap);
  } catch (err) {
    // The bomb guard's own refusal is clean and intended — surface it. Anything else is fflate
    // reporting a malformed archive; replace it wholesale so its raw text never reaches the caller.
    if (err instanceof Error && err.message.startsWith('refusing to inflate')) throw err;
    throw new UnsupportedFormatError('unknown');
  }
}

/**
 * The typed error to raise when an inflated package carries no `xl/workbook.xml`: a `.xlsb` if its
 * binary `xl/workbook.bin` office document is present, otherwise an unrecognised (non-workbook) ZIP.
 */
export function unsupportedWorkbookPart(
  partText: (path: string) => string | undefined,
): UnsupportedFormatError {
  if (partText('xl/workbook.bin') !== undefined) return new UnsupportedFormatError('xlsb');
  return new UnsupportedFormatError('unknown');
}
