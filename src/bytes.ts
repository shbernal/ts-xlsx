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
