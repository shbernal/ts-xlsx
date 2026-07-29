// Front of the read pipe: format detection, the inflate bound both readers share, and the typed
// rejection of everything that is neither an `.xlsx` nor an `.xlsb`.
//
// The reader's first job on an arbitrary blob is to reject what it cannot read with a clear, typed
// error rather than a raw fflate crash. A cheap magic-byte sniff runs *before* the zip layer (so a
// legacy `.xls` compound file or non-ZIP garbage fails fast, before fflate throws its opaque "end of
// central directory" error — which can also leak an absolute path from below), and a malformed ZIP is
// translated after it.
//
// The two outcomes are deliberately different types, because they answer different questions. A blob
// the sniff rejects is the *wrong kind of thing* — {@link UnsupportedFormatError}, branchable on
// `.format`. A `PK`-headed blob the zip layer then chokes on is the right kind of thing we cannot
// unpack — {@link PackageReadError}. Whichever it is, the message states the check that actually ran,
// and callers branch on the type, never on message text.
//
// The `.xlsx` and `.xlsb` serialisations share this whole layer — same container, same bound, same
// rejections — so it is stated once here and neither reader owns it.

import {PackageReadError, UnsupportedFormatError} from './errors.ts';
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
 * Inflate a spreadsheet package (`.xlsx` or `.xlsb` — the container is the same), translating input
 * that is neither into a typed error before or instead of a raw zip failure:
 * - a legacy `.xls` (CFB) blob → {@link UnsupportedFormatError} `'xls'`, caught by the magic sniff so
 *   fflate never runs on it;
 * - a non-ZIP blob → {@link UnsupportedFormatError} `'unknown'`, likewise caught before inflation;
 * - a `PK`-headed blob that fflate then rejects as malformed → {@link PackageReadError}, with the
 *   underlying zip message discarded so no internals (or path) leak.
 *
 * The bounded-inflation guard (a probable zip bomb) already raises {@link PackageReadError} with its
 * own informative message, and is re-thrown unchanged.
 */
export function inflateSpreadsheetPackage(
  data: Uint8Array,
  cap: number,
): Record<string, Uint8Array> {
  const container = sniffContainer(data);
  if (container === 'cfb') throw new UnsupportedFormatError('xls');
  // Not the default message: nothing here has looked for a workbook part, and saying so would point
  // an investigation a layer past the one that actually refused.
  if (container === 'other') {
    throw new UnsupportedFormatError(
      'unknown',
      'not a valid .xlsx package: the input is not a ZIP',
    );
  }
  try {
    return inflatePackage(data, cap);
  } catch (err) {
    // The bomb guard already throws the right type with a better message — surface it. Anything else
    // is fflate reporting a malformed archive: same classification (a ZIP we cannot unpack), but its
    // raw text is replaced wholesale rather than wrapped or attached as `cause`, because it comes from
    // a layer whose strings may name internals — or an absolute path — that must not reach a caller.
    if (err instanceof PackageReadError) throw err;
    throw new PackageReadError(
      'not a readable .xlsx package: the ZIP container is corrupt or truncated and could not be inflated',
    );
  }
}

/**
 * The typed error for an inflated package that carries no `xl/workbook.xml`: a `.xlsb` if its binary
 * `xl/workbook.bin` office document is present, otherwise an unrecognised (non-workbook) ZIP.
 *
 * The `.xlsb` branch takes the caller's own explanation, because whether a binary workbook is
 * readable now depends on *which* entry point was asked: `readXlsx` reads one, the row streamer
 * cannot yet. A single baked-in "not supported" message would be wrong for one of them.
 */
export function unsupportedWorkbookPart(
  partText: (path: string) => string | undefined,
  xlsbMessage: string,
): UnsupportedFormatError {
  if (partText('xl/workbook.bin') !== undefined) {
    return new UnsupportedFormatError('xlsb', xlsbMessage);
  }
  return new UnsupportedFormatError('unknown');
}
