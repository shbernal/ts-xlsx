// The byte primitive every layer needs. It sits beside `errors.ts`, below all of them: the ZIP
// container, the XLSX write stream and the CFB reader all end up holding a list of chunks and
// wanting one buffer, and none of them is the right owner of that for the others.

/**
 * Join byte chunks into one buffer.
 *
 * Both parameters exist to avoid work the caller has already done. `size` is the total length when
 * the caller knows it, as a reader draining a length-prefixed stream does, which skips a pass over
 * the chunks. A lone chunk is handed straight back rather than copied, which is the common case on
 * the inflate path, where a part that fits in one chunk would otherwise be duplicated in memory the
 * moment it is read.
 *
 * The returned buffer therefore aliases the caller's chunk when there is exactly one. Every caller
 * here is a reader assembling bytes it then only reads, so this is sound; a caller that means to
 * mutate the result must copy it.
 */
export function concat(chunks: readonly Uint8Array[], size?: number): Uint8Array {
  const first = chunks[0];
  if (chunks.length === 1 && first !== undefined) return first;
  let total = size;
  if (total === undefined) {
    total = 0;
    for (const chunk of chunks) total += chunk.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

// The base64 alphabet, in index order (RFC 4648 §4).
const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/**
 * Standard base64 of `bytes`, padded with `=`.
 *
 * Spelled out rather than delegated because both platform routes are unavailable here: `Buffer` is
 * a Node global, which is exactly what a browser-safe module may not reach for, and `btoa` is
 * declared deprecated in Node's types (so the `no-deprecated` lint rule rejects it) and takes a
 * binary string rather than bytes anyway. The one caller encodes a 16-byte salt and a 64-byte
 * hash, so the loop below is not on any path where its cost is measurable.
 */
export function toBase64(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i] ?? 0;
    const b = bytes[i + 1] ?? 0;
    const c = bytes[i + 2] ?? 0;
    const triple = (a << 16) | (b << 8) | c;
    const remaining = bytes.length - i;
    // `charAt` rather than an index: `noUncheckedIndexedAccess` types `s[i]` as possibly
    // undefined, and the alphabet is exactly 64 characters wide by construction.
    out += BASE64_ALPHABET.charAt((triple >>> 18) & 63);
    out += BASE64_ALPHABET.charAt((triple >>> 12) & 63);
    out += remaining > 1 ? BASE64_ALPHABET.charAt((triple >>> 6) & 63) : '=';
    out += remaining > 2 ? BASE64_ALPHABET.charAt(triple & 63) : '=';
  }
  return out;
}
