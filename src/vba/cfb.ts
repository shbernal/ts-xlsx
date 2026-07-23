// Reader for the OLE2 / Compound File Binary format ([MS-CFB]).
//
// `vbaProject.bin` is a CFB container (the same "structured storage" behind legacy .doc/.xls). We read
// streams by name and, for the edit-in-place path, reconstruct the whole storage/stream hierarchy so it
// can be re-emitted through the writer with one stream swapped — so this is a deliberate subset of
// [MS-CFB]: header → FAT → directory, plus the mini-FAT for sub-cutoff streams, and the red-black
// sibling tree each storage navigates.
//
// It parses an untrusted blob, so every sector index, chain, and stream size is bounds-checked against
// the file and every chain walk is cycle-guarded. A malformed container fails closed with a
// VbaParseError instead of reading out of bounds, looping forever, or over-allocating.

import type {CfbNode} from './cfb-writer.ts';
import {VbaParseError} from './errors.ts';

interface DirEntry {
  readonly name: string;
  readonly type: number; // 0=empty 1=storage 2=stream 5=root
  readonly startSector: number;
  readonly size: number;
  // Red-black sibling-tree links (entry indices, or a terminal marker for "none"). The reader keeps
  // them so it can walk the tree a host navigates and rebuild the hierarchy, not merely scan names.
  readonly left: number;
  readonly right: number;
  readonly child: number;
}

// Sector values 0xFFFFFFFA..0xFFFFFFFF are reserved markers (DIFSECT/FATSECT/ENDOFCHAIN/FREESECT), not
// data-sector indices; any value at or above this is chain-terminal. Directory-tree links reuse the same
// convention: NOSTREAM (0xFFFFFFFF) and any value at or above the ceiling mean "no such sibling/child".
const MAX_REGULAR_SECTOR = 0xfffffffa;
const NOSTREAM = 0xffffffff;
const TYPE_EMPTY = 0;
const TYPE_STORAGE = 1;
const TYPE_STREAM = 2;
const TYPE_ROOT = 5;
const CFB_SIGNATURE_LO = 0xe011cfd0;
const CFB_SIGNATURE_HI = 0xe11ab1a1;
const DIR_ENTRY_SIZE = 128;

export class CompoundFile {
  readonly #buf: Uint8Array;
  readonly #view: DataView;
  readonly #sectorSize: number;
  readonly #miniSectorSize: number;
  readonly #miniCutoff: number;
  readonly #maxSector: number;
  readonly #fat: number[];
  readonly #miniFat: number[];
  readonly #dir: DirEntry[];
  readonly #miniStream: Uint8Array;

  constructor(buf: Uint8Array) {
    this.#buf = buf;
    this.#view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);

    if (buf.length < 512) throw new VbaParseError('compound file shorter than its 512-byte header');
    if (this.#u32(0) !== CFB_SIGNATURE_LO || this.#u32(4) !== CFB_SIGNATURE_HI) {
      throw new VbaParseError('not a compound file (bad OLE2 signature)');
    }

    const sectorShift = this.#u16(30);
    const miniSectorShift = this.#u16(32);
    // [MS-CFB] fixes these: 512-byte sectors (shift 9) for v3, 4096 (shift 12) for v4; mini shift 6.
    if (sectorShift !== 9 && sectorShift !== 12) {
      throw new VbaParseError(`unsupported sector shift ${sectorShift}`);
    }
    if (miniSectorShift !== 6)
      throw new VbaParseError(`unsupported mini-sector shift ${miniSectorShift}`);
    this.#sectorSize = 1 << sectorShift;
    this.#miniSectorSize = 1 << miniSectorShift;
    this.#maxSector = Math.floor(buf.length / this.#sectorSize);

    const numFatSectors = this.#u32(44);
    const firstDirSector = this.#u32(48);
    this.#miniCutoff = this.#u32(56);
    const firstMiniFatSector = this.#u32(60);
    const firstDifatSector = this.#u32(68);
    const numDifatSectors = this.#u32(72);

    const fatSectorIds = this.#readDifat(numFatSectors, firstDifatSector, numDifatSectors);
    this.#fat = this.#readFat(fatSectorIds);
    this.#miniFat = this.#readChainValues(firstMiniFatSector);
    this.#dir = this.#readDirectory(firstDirSector);

    const root = this.#dir.find((e) => e.type === TYPE_ROOT);
    if (!root) throw new VbaParseError('compound file has no root storage entry');
    // The mini-stream lives in the regular FAT, addressed from the root entry's start sector.
    this.#miniStream = this.#readViaFat(root.startSector, root.size);
  }

  /** List every stream/storage name in the directory (order as stored). */
  names(): string[] {
    return this.#dir
      .filter((e) => e.type === TYPE_STREAM || e.type === TYPE_STORAGE)
      .map((e) => e.name);
  }

  /** Read a stream's raw bytes by exact entry name, or `undefined` if absent. */
  readStream(name: string): Uint8Array | undefined {
    const entry = this.#dir.find((e) => e.type === TYPE_STREAM && e.name === name);
    if (!entry) return undefined;
    return this.#readEntryData(entry);
  }

  /**
   * Reconstruct the container's top-level children as the writer's node shape, recursing into every
   * storage — so a caller can swap one stream and re-emit the whole hierarchy with {@link writeCompoundFile}.
   * Walks the red-black sibling tree each storage navigates (not the linear directory scan), so any part
   * a host reaches is carried through. Cycle- and bounds-guarded like every other chain walk here.
   */
  tree(): CfbNode[] {
    const rootIdx = this.#dir.findIndex((e) => e.type === TYPE_ROOT);
    if (rootIdx < 0) throw new VbaParseError('compound file has no root storage entry');
    const root = this.#dir[rootIdx] as DirEntry;
    return this.#buildSiblings(root.child, new Set([rootIdx]));
  }

  #buildSiblings(firstChild: number, seen: Set<number>): CfbNode[] {
    const nodes: CfbNode[] = [];
    const walk = (idx: number): void => {
      if (idx >= MAX_REGULAR_SECTOR) return; // NOSTREAM / terminal marker → no such sibling
      if (idx >= this.#dir.length) throw new VbaParseError('directory sibling index out of range');
      if (seen.has(idx)) throw new VbaParseError('cycle in directory sibling tree');
      seen.add(idx);
      const e = this.#dir[idx] as DirEntry;
      if (e.type === TYPE_EMPTY) throw new VbaParseError('directory tree links an empty entry');
      walk(e.left);
      if (e.type === TYPE_STORAGE) {
        nodes.push({name: e.name, children: this.#buildSiblings(e.child, seen)});
      } else {
        nodes.push({name: e.name, data: this.#readEntryData(e)});
      }
      walk(e.right);
    };
    walk(firstChild);
    return nodes;
  }

  #readEntryData(entry: DirEntry): Uint8Array {
    if (entry.size >= this.#miniCutoff) return this.#readViaFat(entry.startSector, entry.size);
    return this.#readViaMiniFat(entry.startSector, entry.size);
  }

  #readDifat(numFatSectors: number, firstDifat: number, numDifat: number): number[] {
    const ids: number[] = [];
    // The first 109 FAT-sector pointers live in the header; the rest chain through DIFAT sectors.
    for (let i = 0; i < 109 && ids.length < numFatSectors; i++) {
      const v = this.#u32(76 + i * 4);
      if (v >= MAX_REGULAR_SECTOR) break;
      ids.push(v);
    }
    const seen = new Set<number>();
    let sector = firstDifat;
    const perSector = this.#sectorSize / 4 - 1;
    for (let s = 0; s < numDifat && sector < MAX_REGULAR_SECTOR; s++) {
      if (seen.has(sector)) throw new VbaParseError('cycle in DIFAT sector chain');
      seen.add(sector);
      const base = this.#dataSectorOffset(sector);
      for (let i = 0; i < perSector; i++) {
        const v = this.#u32(base + i * 4);
        if (v < MAX_REGULAR_SECTOR) ids.push(v);
      }
      sector = this.#u32(base + perSector * 4);
    }
    return ids;
  }

  #readFat(fatSectorIds: number[]): number[] {
    const fat: number[] = [];
    const perSector = this.#sectorSize / 4;
    for (const sid of fatSectorIds) {
      const base = this.#dataSectorOffset(sid);
      for (let i = 0; i < perSector; i++) fat.push(this.#u32(base + i * 4));
    }
    return fat;
  }

  #readChainValues(firstSector: number): number[] {
    // Every uint32 in a sector chain (used for the mini-FAT), cycle-guarded.
    const values: number[] = [];
    const perSector = this.#sectorSize / 4;
    const seen = new Set<number>();
    let sector = firstSector;
    while (sector < MAX_REGULAR_SECTOR) {
      if (seen.has(sector)) throw new VbaParseError('cycle in mini-FAT sector chain');
      seen.add(sector);
      const base = this.#dataSectorOffset(sector);
      for (let i = 0; i < perSector; i++) values.push(this.#u32(base + i * 4));
      sector = this.#nextInFat(sector);
    }
    return values;
  }

  #readDirectory(firstDirSector: number): DirEntry[] {
    const raw = this.#readChainFull(firstDirSector);
    const entries: DirEntry[] = [];
    // Empty slots are kept as placeholders (not skipped) so array indices stay equal to the on-disk
    // directory-entry ids the sibling-tree links reference — the tree walk in #buildSiblings needs them.
    for (let off = 0; off + DIR_ENTRY_SIZE <= raw.length; off += DIR_ENTRY_SIZE) {
      const type = raw[off + 66] as number;
      if (
        type !== TYPE_EMPTY &&
        type !== TYPE_STORAGE &&
        type !== TYPE_STREAM &&
        type !== TYPE_ROOT
      ) {
        throw new VbaParseError(`directory entry has invalid object type ${type}`);
      }
      const left = readU32(raw, off + 68);
      const right = readU32(raw, off + 72);
      const child = readU32(raw, off + 76);
      if (type === TYPE_EMPTY) {
        entries.push({
          name: '',
          type,
          startSector: 0,
          size: 0,
          left: NOSTREAM,
          right: NOSTREAM,
          child: NOSTREAM,
        });
        continue;
      }
      const nameLen = readU16(raw, off + 64);
      if (nameLen > 64 || nameLen % 2 !== 0) {
        throw new VbaParseError(`directory entry has invalid name length ${nameLen}`);
      }
      const name = decodeUtf16le(raw.subarray(off, off + Math.max(0, nameLen - 2)));
      const startSector = readU32(raw, off + 116);
      const size = readU32(raw, off + 120); // low 32 bits — ample for a VBA project
      entries.push({name, type, startSector, size, left, right, child});
    }
    return entries;
  }

  #readChainFull(startSector: number): Uint8Array {
    // Follow a FAT chain to its end, collecting whole sectors (used for the directory, whose byte
    // length is not declared). Cycle-guarded and bounded by the sector count.
    const chunks: Uint8Array[] = [];
    const seen = new Set<number>();
    let sector = startSector;
    while (sector < MAX_REGULAR_SECTOR) {
      if (seen.has(sector)) throw new VbaParseError('cycle in FAT sector chain');
      seen.add(sector);
      const base = this.#dataSectorOffset(sector);
      chunks.push(this.#buf.subarray(base, base + this.#sectorSize));
      sector = this.#nextInFat(sector);
    }
    return concat(chunks);
  }

  #readViaFat(startSector: number, size: number): Uint8Array {
    const chunks: Uint8Array[] = [];
    const seen = new Set<number>();
    let sector = startSector;
    let remaining = size;
    while (sector < MAX_REGULAR_SECTOR && remaining > 0) {
      if (seen.has(sector)) throw new VbaParseError('cycle in stream FAT chain');
      seen.add(sector);
      const base = this.#dataSectorOffset(sector);
      const take = Math.min(this.#sectorSize, remaining);
      chunks.push(this.#buf.subarray(base, base + take));
      remaining -= take;
      sector = this.#nextInFat(sector);
    }
    return concat(chunks);
  }

  #readViaMiniFat(startSector: number, size: number): Uint8Array {
    const chunks: Uint8Array[] = [];
    const seen = new Set<number>();
    let sector = startSector;
    let remaining = size;
    while (sector < MAX_REGULAR_SECTOR && remaining > 0) {
      if (seen.has(sector)) throw new VbaParseError('cycle in stream mini-FAT chain');
      seen.add(sector);
      const base = sector * this.#miniSectorSize;
      const take = Math.min(this.#miniSectorSize, remaining);
      if (base + take > this.#miniStream.length) {
        throw new VbaParseError('mini-stream sector runs past end of the mini stream');
      }
      chunks.push(this.#miniStream.subarray(base, base + take));
      remaining -= take;
      if (sector >= this.#miniFat.length) throw new VbaParseError('mini-FAT index out of range');
      sector = this.#miniFat[sector] as number;
    }
    return concat(chunks);
  }

  #nextInFat(sector: number): number {
    if (sector >= this.#fat.length) throw new VbaParseError('FAT index out of range');
    return this.#fat[sector] as number;
  }

  #dataSectorOffset(sector: number): number {
    if (sector >= this.#maxSector) {
      throw new VbaParseError(`sector ${sector} is out of range (file has ${this.#maxSector})`);
    }
    const base = (sector + 1) * this.#sectorSize;
    if (base + this.#sectorSize > this.#buf.length) {
      throw new VbaParseError(`sector ${sector} runs past end of file`);
    }
    return base;
  }

  #u16(at: number): number {
    return this.#view.getUint16(at, true);
  }
  #u32(at: number): number {
    return this.#view.getUint32(at, true);
  }
}

function readU16(buf: Uint8Array, at: number): number {
  return (buf[at] as number) | ((buf[at + 1] as number) << 8);
}
function readU32(buf: Uint8Array, at: number): number {
  return (
    ((buf[at] as number) |
      ((buf[at + 1] as number) << 8) |
      ((buf[at + 2] as number) << 16) |
      ((buf[at + 3] as number) << 24)) >>>
    0
  );
}
function decodeUtf16le(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    s += String.fromCharCode((bytes[i] as number) | ((bytes[i + 1] as number) << 8));
  }
  return s;
}
function concat(chunks: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const c of chunks) total += c.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}
