// The one knob every reader takes, declared apart from any of them.
//
// `readXlsx`, `readXlsb`, and the row streamer all inflate the same OPC container under the same
// bound, so the option bag belongs to none of them individually, and a reader that dispatches to
// another (the `.xlsx` entry point hands a binary package to the `.xlsb` codec) cannot import it from
// its peer without the two modules importing each other.
//
// It is published from `/core`, not from a codec. It used to be `ReadXlsxOptions`, exported only from
// `/xlsx`, which left a `/xlsb` consumer unable to name the option type of the one function that
// subpath publishes without importing the very codec `/xlsb` exists to let them avoid. Nothing caught
// that: the shape is structural, so it typechecks either way, and `check-entries.ts` could not, since
// its disjointness rule *forbids* publishing one name from two subpaths. The old name was wrong on its
// face too: this bounds the OPC container, which all three readers share, not the XML codec.

export interface ReadPackageOptions {
  /**
   * Maximum total uncompressed output, in bytes, produced while inflating the package.
   * The bound is enforced by a running counter as bytes are decompressed, never read from
   * the archive's (untrusted, forgeable) size headers, so a zip bomb that lies about its
   * uncompressed size is rejected all the same. Defaults to 512 MiB.
   */
  readonly maxUncompressedBytes?: number;
}

export const DEFAULT_MAX_UNCOMPRESSED = 512 * 1024 * 1024;
