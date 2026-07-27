// Typed errors the reader raises when it is handed something that is not a readable OOXML `.xlsx`
// package. A spreadsheet library parses untrusted files, so the failure a caller sees must be a clear,
// programmatically-branchable signal — not a raw zip-internals string (which is opaque, and can leak an
// absolute filesystem path from the layer below). See the spec
// `docs/knowledge/specs/unsupported-input-format-typed-error.md`.

/**
 * Which unsupported input the reader recognised:
 * - `'xls'` — a legacy BIFF `.xls` (an OLE2/CFB compound file), detected by its magic bytes.
 * - `'xlsb'` — a binary BIFF12 `.xlsb`: the same OPC/ZIP container as `.xlsx`, but its office document
 *   is `xl/workbook.bin` rather than `xl/workbook.xml`. `readXlsx`/`readXlsb` read one; the entry points
 *   that cannot yet (the row streamer) report it under this format with their own message.
 * - `'unknown'` — not a recognised spreadsheet at all: not a ZIP, or a ZIP carrying no OOXML workbook
 *   part (nor a `.xlsb` binary one).
 */
export type UnsupportedFormat = 'xls' | 'xlsb' | 'unknown';

const DEFAULT_MESSAGE: Record<UnsupportedFormat, string> = {
  xls: 'the legacy .xls binary format (BIFF/OLE2 compound file) is not supported; only OOXML .xlsx is read',
  xlsb: 'this entry point does not read the binary .xlsb format (BIFF12)',
  unknown: 'not a valid .xlsx package: no OOXML workbook part was found',
};

/**
 * Thrown when input is not a readable `.xlsx` package. The single {@link format} field is the branch a
 * caller keys on (rather than a subclass per format), so a `catch` can distinguish a legacy `.xls`, a
 * binary `.xlsb`, and an unrecognised blob without string-matching the message.
 *
 * The message never carries a filesystem path or the underlying zip library's internals — the whole
 * point of the type is that the classification, not a leaked lower-layer string, is what the caller sees.
 */
export class UnsupportedFormatError extends Error {
  override readonly name = 'UnsupportedFormatError';
  readonly format: UnsupportedFormat;

  constructor(format: UnsupportedFormat, message: string = DEFAULT_MESSAGE[format]) {
    super(message);
    this.format = format;
  }
}
