// The one knob every reader takes, declared apart from any of them.
//
// `readXlsx`, `readXlsb`, and the row streamer all inflate the same OPC container under the same
// bound, so the option bag belongs to none of them individually — and a reader that dispatches to
// another (the `.xlsx` entry point hands a binary package to the `.xlsb` codec) cannot import it from
// its peer without the two modules importing each other.

export interface ReadXlsxOptions {
  /**
   * Maximum total uncompressed output, in bytes, produced while inflating the package.
   * The bound is enforced by a running counter as bytes are decompressed — never read from
   * the archive's (untrusted, forgeable) size headers — so a zip bomb that lies about its
   * uncompressed size is rejected all the same. Defaults to 512 MiB.
   */
  readonly maxUncompressedBytes?: number;
}

export const DEFAULT_MAX_UNCOMPRESSED = 512 * 1024 * 1024;
