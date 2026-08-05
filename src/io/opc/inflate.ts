// Bounded, streaming inflation of an OPC (`.xlsx`) zip package.
//
// Decompression is the reader's first hostile-input surface: a "zip bomb" ships a few
// kilobytes that inflate to gigabytes. The zip's own size headers cannot be trusted to
// bound this — they are attacker-controlled. A header that declares a *large* size makes a
// naïve reader preallocate that much (an amplifier: tiny input, huge allocation); a header
// that lies *small* makes a size-preallocating inflater silently truncate real data. So we
// consult the declared sizes for nothing. Instead the compressed archive is fed to fflate's
// streaming unzip in slices, the decompressor grows its output from the bytes it actually
// produces, and a running counter of real output aborts the moment it crosses the cap.
//
// Because DEFLATE cannot expand input by more than ~1032:1, feeding at most INPUT_SLICE
// compressed bytes before each counter check bounds the worst-case overshoot past the cap
// to one slice's expansion — not the whole (possibly enormous) stream.

import {type FlateError, Unzip, type UnzipFile, UnzipInflate} from 'fflate';
import {PackageReadError} from './errors.ts';

// Compressed input is pushed in slices this size so decompressed output arrives in
// increments the running counter can check. Small enough that the worst-case overshoot
// (slice × DEFLATE's ~1032:1 ceiling ≈ 16 MiB) is negligible against a sane cap; large
// enough that a legitimate multi-hundred-megabyte package still streams in cheaply.
const INPUT_SLICE = 1 << 14;

/**
 * Inflate every part of an `.xlsx` zip package, enforcing a hard ceiling on total
 * decompressed output.
 *
 * @param data The raw zip bytes.
 * @param cap  Maximum total uncompressed output, in bytes, across all parts. Enforced
 *   against bytes actually produced, never against the archive's declared sizes.
 * @returns A map of part path to inflated bytes.
 * @throws {PackageReadError} if inflation would exceed `cap` — a probable zip bomb.
 * @throws {Error} raised by the zip layer if the archive is malformed or a part uses an unsupported
 *   compression method. Callers reach this through {@link inflateSpreadsheetPackage}, which is where
 *   those are classified; the raw error is deliberately not re-typed here, because its text is the
 *   one thing that must not reach a caller unfiltered.
 */
export function inflatePackage(data: Uint8Array, cap: number): Record<string, Uint8Array> {
  const files: Record<string, Uint8Array> = {};
  let total = 0;
  let failure: Error | undefined;

  const unzip = new Unzip((file: UnzipFile) => {
    const chunks: Uint8Array[] = [];
    let size = 0;
    file.ondata = (error: FlateError | null, chunk: Uint8Array, final: boolean): void => {
      if (failure) return;
      if (error) {
        failure = error;
        return;
      }
      total += chunk.length;
      if (total > cap) {
        failure = new PackageReadError(
          `refusing to inflate: uncompressed output exceeds ${cap} bytes (possible zip bomb)`,
        );
        return;
      }
      chunks.push(chunk);
      size += chunk.length;
      if (final) files[file.name] = join(chunks, size);
    };
    // `start()` reports an unsupported compression method through `ondata` above, then
    // throws trying to build the missing decoder. Keep the reported error (its message
    // names the method) and swallow the raw follow-on throw.
    try {
      file.start();
    } catch (error) {
      failure ??= error instanceof Error ? error : new Error(String(error));
    }
  });
  unzip.register(UnzipInflate);

  let offset = 0;
  do {
    const end = Math.min(offset + INPUT_SLICE, data.length);
    unzip.push(data.subarray(offset, end), end === data.length);
    offset = end;
  } while (offset < data.length && !failure);

  if (failure) throw failure;
  return files;
}

function join(chunks: Uint8Array[], size: number): Uint8Array {
  if (chunks.length === 1) return chunks[0] as Uint8Array;
  const out = new Uint8Array(size);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}
