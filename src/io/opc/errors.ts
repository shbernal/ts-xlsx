// Typed errors the reader raises when it is handed something that is not a readable OOXML `.xlsx`
// package. A spreadsheet library parses untrusted files, so the failure a caller sees must be a clear,
// programmatically-branchable signal — not a raw zip-internals string (which is opaque, and can leak an
// absolute filesystem path from the layer below). See the spec
// `docs/knowledge/specs/unsupported-input-format-typed-error.md`.

import {XlsxError} from '../../errors.ts';

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
 *
 * {@link format} stays the branch for *which* unsupported input this was; the inherited
 * {@link XlsxError.code} answers the coarser question of what kind of failure it is.
 */
export class UnsupportedFormatError extends XlsxError {
  override readonly name = 'UnsupportedFormatError';
  override readonly code = 'unsupported-format';
  readonly format: UnsupportedFormat;

  constructor(
    format: UnsupportedFormat,
    message: string = DEFAULT_MESSAGE[format],
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.format = format;
  }
}

/**
 * Thrown when the input *is* a ZIP container but reading it is refused — today, when inflation would
 * push total uncompressed output past the caller's bound, which is how a zip bomb presents.
 *
 * The neighbouring {@link UnsupportedFormatError} says the input is a different *kind* of thing; this
 * one says it is the right kind and we will not (or cannot) unpack it. Keeping them apart is what
 * lets a caller answer "should I try another reader, or reject this file?" — and it is what replaced
 * the message-prefix match this refusal used to be recognised by.
 *
 * A zip-library failure underneath is *not* re-thrown as this type: its text can name internals, so
 * it is still classified as an unrecognised package rather than surfaced. Should that change, the
 * original belongs on `cause`, never folded into the message.
 */
export class PackageReadError extends XlsxError {
  override readonly name = 'PackageReadError';
  override readonly code = 'malformed-input';
}
