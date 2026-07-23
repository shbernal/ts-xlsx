// MS-OVBA §2.4.1 — compression and decompression of a "CompressedContainer".
//
// VBA module source and the project `dir` stream are stored in Office's own run-length compression,
// NOT deflate. A container is a 0x01 signature byte followed by one or more chunks; each chunk
// decompresses to at most 4096 bytes and is either a raw 4096-byte copy or a stream of literal/copy
// tokens. Reference: [MS-OVBA] 2.4.1.3.6 (decompressing a CompressedContainer), 2.4.1.3.19.3 (the
// CopyToken bit-packing), and 2.4.1.3.7 (compressing a chunk).
//
// The decompressor is a hostile-input parser: the container comes from an untrusted file, so every
// length and back-reference is bounds-checked and the total output is capped. A malformed container
// fails closed with a VbaParseError rather than over-allocating, looping, or reading out of bounds. The
// compressor is the authoring inverse — it is fed our own bytes, and its output re-expands to the input
// byte-for-byte (the round-trip is the correctness contract).

import {VbaParseError} from './errors.ts';

// A decompressed chunk covers at most 4096 bytes; both directions honour this window ([MS-OVBA]
// 2.4.1.3.6). The chunk header's bits 12-14 carry a fixed 0b011 signature, bit 15 the compressed flag.
const MAX_CHUNK_DECOMPRESSED = 4096;
const CHUNK_SIGNATURE = 0b011 << 12;
const CHUNK_COMPRESSED_FLAG = 0x8000;

// A single VBA project is well under a megabyte; 64 MiB is far above any legitimate container yet
// bounds a decompression bomb (a small container that expands without limit) to a survivable size.
const DEFAULT_MAX_OUTPUT = 64 * 1024 * 1024;

/**
 * Decompress an MS-OVBA CompressedContainer beginning at `start` in `buf`.
 * @param maxOutput hard ceiling on decompressed bytes; exceeding it throws (bomb guard).
 */
export function decompressContainer(
  buf: Uint8Array,
  start = 0,
  maxOutput = DEFAULT_MAX_OUTPUT,
): Uint8Array {
  if (start >= buf.length) {
    throw new VbaParseError(
      `compressed container starts past end of stream (${start} >= ${buf.length})`,
    );
  }
  if (buf[start] !== 0x01) {
    throw new VbaParseError(
      `compressed container must begin with a 0x01 signature byte, found 0x${(buf[start] ?? 0).toString(16)}`,
    );
  }

  const out: number[] = [];
  let pos = start + 1;

  while (pos + 2 <= buf.length) {
    const header = readU16(buf, pos);
    pos += 2;

    // Bits 0-11: (chunk data size - 1). Bit 15: compressed flag. Bits 12-14: the fixed 0b011 signature.
    const chunkDataSize = (header & 0x0fff) + 1;
    const compressed = (header & 0x8000) !== 0;
    if (((header >> 12) & 0x7) !== 0b011) {
      throw new VbaParseError(`chunk header has a bad 0b011 signature (0x${header.toString(16)})`);
    }
    const chunkEnd = pos + chunkDataSize;
    if (chunkEnd > buf.length) {
      throw new VbaParseError(`chunk data (${chunkDataSize} bytes) runs past end of stream`);
    }

    if (!compressed) {
      // A raw chunk carries its bytes verbatim (Excel emits one only when compression would expand).
      for (let i = pos; i < chunkEnd; i++) out.push(buf[i] as number);
      guardOutput(out.length, maxOutput);
      pos = chunkEnd;
      continue;
    }

    const chunkStart = out.length; // decompressed offset where this chunk began
    while (pos < chunkEnd) {
      const flagByte = buf[pos++] as number;
      for (let bit = 0; bit < 8 && pos < chunkEnd; bit++) {
        const isCopy = (flagByte >> bit) & 1;
        if (!isCopy) {
          out.push(buf[pos++] as number);
          guardOutput(out.length, maxOutput);
          continue;
        }
        if (pos + 2 > chunkEnd) {
          throw new VbaParseError('copy token truncated at chunk end');
        }
        const token = readU16(buf, pos);
        pos += 2;
        const {lengthMask, bitCount} = copyTokenHelp(out.length - chunkStart);
        const length = (token & lengthMask) + 3;
        const copyOffset = (token >> (16 - bitCount)) + 1;
        const src = out.length - copyOffset;
        if (src < chunkStart) {
          // A back-reference may only reach data emitted since this chunk began ([MS-OVBA] 2.4.1.3.19).
          throw new VbaParseError('copy token references before the start of its chunk');
        }
        // Byte-by-byte so overlapping runs (run-length expansion) grow correctly.
        for (let i = 0; i < length; i++) out.push(out[src + i] as number);
        guardOutput(out.length, maxOutput);
      }
    }
    pos = chunkEnd;
  }

  return Uint8Array.from(out);
}

/**
 * Compress `data` into an MS-OVBA CompressedContainer — the inverse of {@link decompressContainer}.
 * Every 4096-decompressed-byte window is emitted as a compressed chunk of literal and copy tokens, or
 * stored verbatim when compression would not shrink it (so the encoded chunk never exceeds the 12-bit
 * size field). The result re-expands to `data` byte-for-byte.
 */
export function compressContainer(data: Uint8Array): Uint8Array {
  const out: number[] = [0x01]; // container signature; an empty input yields just this byte
  for (let start = 0; start < data.length; start += MAX_CHUNK_DECOMPRESSED) {
    const chunk = data.subarray(start, Math.min(start + MAX_CHUNK_DECOMPRESSED, data.length));
    const tokens = compressChunk(chunk);
    // Prefer the token stream only when it is strictly smaller; otherwise store the chunk raw. Both
    // encode their exact length in the header, so the decompressor reconstructs the window either way.
    const compressed = tokens.length < chunk.length;
    const body = compressed ? tokens : chunk;
    const header =
      (compressed ? CHUNK_COMPRESSED_FLAG : 0) | CHUNK_SIGNATURE | ((body.length - 1) & 0x0fff);
    out.push(header & 0xff, (header >> 8) & 0xff);
    for (const b of body) out.push(b);
  }
  return Uint8Array.from(out);
}

// Encode one decompressed chunk (≤ 4096 bytes) as a sequence of MS-OVBA token groups: a flag byte whose
// bits mark the next up-to-8 tokens as literal (0) or copy (1). A copy token replaces a run of 3+ bytes
// that recurs earlier in the *same* chunk; matches may overlap the current position (run-length growth),
// which the decompressor reproduces byte-by-byte. The bit split between the offset and length fields
// widens as the chunk fills, exactly as the decoder computes it, so both agree on every token's shape.
function compressChunk(chunk: Uint8Array): number[] {
  const tokens: number[] = [];
  let pos = 0;
  while (pos < chunk.length) {
    const flagIndex = tokens.length;
    tokens.push(0);
    let flags = 0;
    for (let bit = 0; bit < 8 && pos < chunk.length; bit++) {
      const {lengthMask, bitCount} = copyTokenHelp(pos);
      const maxLength = lengthMask + 3;
      const windowStart = Math.max(0, pos - (1 << bitCount));

      let bestLength = 0;
      let bestOffset = 0;
      // Scan nearest-first so equal-length matches keep the smallest offset (a marginally cheaper token).
      for (let cand = pos - 1; cand >= windowStart; cand--) {
        let len = 0;
        while (
          len < maxLength &&
          pos + len < chunk.length &&
          chunk[cand + len] === chunk[pos + len]
        ) {
          len++;
        }
        if (len > bestLength) {
          bestLength = len;
          bestOffset = pos - cand;
          if (bestLength === maxLength) break; // cannot improve
        }
      }

      if (bestLength >= 3) {
        const token = ((bestOffset - 1) << (16 - bitCount)) | (bestLength - 3);
        tokens.push(token & 0xff, (token >> 8) & 0xff);
        flags |= 1 << bit;
        pos += bestLength;
      } else {
        tokens.push(chunk[pos] as number);
        pos++;
      }
    }
    tokens[flagIndex] = flags;
  }
  return tokens;
}

function guardOutput(size: number, maxOutput: number): void {
  if (size > maxOutput) {
    throw new VbaParseError(
      `decompressed output exceeds the ${maxOutput}-byte ceiling (possible bomb)`,
    );
  }
}

/**
 * Bit split for a CopyToken given how many bytes have been emitted since the current chunk began
 * ([MS-OVBA] 2.4.1.3.19.3): the offset field grows and the length field shrinks as the chunk fills.
 */
function copyTokenHelp(decompressedSoFar: number): {lengthMask: number; bitCount: number} {
  const bitCount = Math.max(Math.ceil(Math.log2(Math.max(decompressedSoFar, 1))), 4);
  const lengthMask = 0xffff >> bitCount;
  return {lengthMask, bitCount};
}

function readU16(buf: Uint8Array, at: number): number {
  return (buf[at] as number) | ((buf[at + 1] as number) << 8);
}
