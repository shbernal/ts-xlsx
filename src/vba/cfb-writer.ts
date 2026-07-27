// Writer for the OLE2 / Compound File Binary format ([MS-CFB]) — the encode counterpart to cfb.ts.
//
// Produces a v3 (512-byte sector) container from a hierarchy of storages and streams, the substrate a
// synthesized vbaProject.bin is built on (its modules live inside a `VBA` storage, with `PROJECT` and
// `PROJECTwm` at the root). Streams below the 4096-byte mini cutoff are packed into the mini stream and
// chained through the mini-FAT; larger streams take whole regular sectors. Each storage's children are
// emitted as a name-ordered balanced binary tree ([MS-CFB] 2.6.4), so a host that *navigates* the tree
// (Excel) reaches every entry — not only a linear scanner like this library's own reader.
//
// Unlike the reader, this is not a hostile-input path: we are the producer. It still validates its
// contract (name length, sibling-name uniqueness, size bound) and fails closed with VbaAuthorError,
// because a silently malformed container would surface far downstream as an unopenable workbook.

import {VbaAuthorError} from './errors.ts';

export interface CfbStream {
  /** The exact directory-entry name. At most 31 UTF-16 code units ([MS-CFB] name limit). */
  readonly name: string;
  readonly data: Uint8Array;
}

export interface CfbStorage {
  readonly name: string;
  readonly children: readonly CfbNode[];
}

export type CfbNode = CfbStream | CfbStorage;

const SECTOR = 512;
const MINI_SECTOR = 64;
const MINI_CUTOFF = 4096;
const DIR_ENTRY_SIZE = 128;
const ENTRIES_PER_DIR_SECTOR = SECTOR / DIR_ENTRY_SIZE; // 4
const FAT_ENTRIES_PER_SECTOR = SECTOR / 4; // 128
const DIFAT_HEADER_SLOTS = 109; // FAT-sector pointers that fit in the header before DIFAT sectors
const MAX_NAME_CHARS = 31; // 32 UTF-16 code units incl. the NUL terminator

// Sector chain markers ([MS-CFB] 2.2).
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;
const NOSTREAM = 0xffffffff;

// Object types ([MS-CFB] 2.6.1).
const TYPE_STORAGE = 1;
const TYPE_STREAM = 2;
const TYPE_ROOT = 5;
const COLOR_BLACK = 1;

function isStream(node: CfbNode): node is CfbStream {
  return 'data' in node;
}

interface DirEntry {
  /**
   * Position in the directory stream. Sibling and child links are stored as these indices, so an
   * entry carries its own — the alternative is looking the object back up in `entries`, which under
   * `noUncheckedIndexedAccess` yields `DirEntry | undefined` at every use site.
   */
  readonly index: number;
  name: string;
  type: number;
  startSector: number;
  size: number;
  left: number;
  right: number;
  child: number;
}

/** A stream too large for the mini stream: it owns whole sectors, placed during layout. */
interface BigStream {
  entry: DirEntry;
  data: Uint8Array;
  sectors: number;
}

/**
 * Encode a hierarchy of storages and streams into a v3 compound file. The Root Entry is synthesized
 * automatically; `root` is its top-level children. Every stream becomes a directory entry reachable both
 * by linear scan and by tree navigation.
 *
 * @throws {VbaAuthorError} if any name is empty or exceeds 31 characters, sibling names collide, or the
 *   project is so large it would need more than 109 FAT sectors (~7 MB — far beyond any real project).
 */
export function writeCompoundFile(root: readonly CfbNode[]): Uint8Array {
  const entries: DirEntry[] = [];
  const addEntry = (name: string, type: number, startSector: number): DirEntry => {
    const entry: DirEntry = {
      index: entries.length,
      name,
      type,
      startSector,
      size: 0,
      left: NOSTREAM,
      right: NOSTREAM,
      child: NOSTREAM,
    };
    entries.push(entry);
    return entry;
  };

  const rootEntry = addEntry('Root Entry', TYPE_ROOT, ENDOFCHAIN);

  // Sub-cutoff stream bytes accumulate into the mini stream (chained in the mini-FAT); larger streams
  // are laid out later directly in the regular FAT. Depth-first walk fixes a deterministic layout.
  const miniBytes: number[] = [];
  const miniFat: number[] = [];
  const bigStreams: BigStream[] = [];

  // `siblings` is the parent's child list itself rather than its index, so a child is appended to an
  // array we hold — no lookup that could come back empty.
  const addNode = (node: CfbNode, siblings: DirEntry[]): void => {
    if (isStream(node)) {
      const entry = addEntry(node.name, TYPE_STREAM, ENDOFCHAIN);
      entry.size = node.data.length;
      siblings.push(entry);
      if (node.data.length === 0) {
        // an empty stream owns no sectors
      } else if (node.data.length >= MINI_CUTOFF) {
        bigStreams.push({entry, data: node.data, sectors: Math.ceil(node.data.length / SECTOR)});
      } else {
        const startMini = miniBytes.length / MINI_SECTOR;
        const numMini = Math.ceil(node.data.length / MINI_SECTOR);
        for (let k = 0; k < numMini; k++)
          miniFat.push(k < numMini - 1 ? startMini + k + 1 : ENDOFCHAIN);
        miniBytes.push(
          ...node.data,
          ...new Array(numMini * MINI_SECTOR - node.data.length).fill(0),
        );
        entry.startSector = startMini;
      }
    } else {
      const entry = addEntry(node.name, TYPE_STORAGE, 0);
      siblings.push(entry);
      const kids: DirEntry[] = [];
      for (const c of node.children) addNode(c, kids);
      // Each storage (Root included) links its children as a balanced search tree the host navigates.
      linkChildren(entry, kids);
    }
  };

  validateSiblingNames('Root Entry', root);
  const rootKids: DirEntry[] = [];
  for (const n of root) addNode(n, rootKids);
  linkChildren(rootEntry, rootKids);

  // ── Sector layout ─────────────────────────────────────────────────────────────────────────────────
  // Physical order: directory, mini-FAT, mini stream, each big stream, FAT, DIFAT. Region starts are
  // assigned first so chains can reference them; the FAT is filled last, once every sector is placed.
  const dirSectors = Math.ceil(entries.length / ENTRIES_PER_DIR_SECTOR);
  const miniFatSectors =
    miniFat.length > 0 ? Math.ceil((miniFat.length * 4) / FAT_ENTRIES_PER_SECTOR) : 0;
  const miniStreamSectors = Math.ceil(miniBytes.length / SECTOR);
  const baseSectors =
    dirSectors +
    miniFatSectors +
    miniStreamSectors +
    bigStreams.reduce((total, big) => total + big.sectors, 0);

  let fatSectors = 0;
  let difatSectors = 0;
  for (;;) {
    const total = baseSectors + fatSectors + difatSectors;
    const needFat = Math.ceil(total / FAT_ENTRIES_PER_SECTOR);
    // Each DIFAT sector holds 127 FAT pointers + a next-DIFAT pointer; the first 109 live in the header.
    const needDifat =
      needFat > DIFAT_HEADER_SLOTS
        ? Math.ceil((needFat - DIFAT_HEADER_SLOTS) / (FAT_ENTRIES_PER_SECTOR - 1))
        : 0;
    if (needFat === fatSectors && needDifat === difatSectors) break;
    fatSectors = needFat;
    difatSectors = needDifat;
  }
  if (fatSectors > DIFAT_HEADER_SLOTS) {
    throw new VbaAuthorError(
      `project needs ${fatSectors} FAT sectors, exceeding the ${DIFAT_HEADER_SLOTS}-sector single-header bound`,
    );
  }

  let cursor = 0;
  const dirStart = cursor;
  cursor += dirSectors;
  const miniFatStart = miniFatSectors > 0 ? cursor : ENDOFCHAIN;
  cursor += miniFatSectors;
  const miniStreamStart = miniStreamSectors > 0 ? cursor : ENDOFCHAIN;
  cursor += miniStreamSectors;
  for (const big of bigStreams) {
    big.entry.startSector = cursor;
    cursor += big.sectors;
  }
  const fatStart = cursor;
  cursor += fatSectors;
  const difatStart = difatSectors > 0 ? cursor : ENDOFCHAIN;
  cursor += difatSectors;
  const totalSectors = cursor;

  rootEntry.startSector = miniStreamStart;
  rootEntry.size = miniBytes.length;

  // ── FAT ─────────────────────────────────────────────────────────────────────────────────────────
  const fat = new Array<number>(fatSectors * FAT_ENTRIES_PER_SECTOR).fill(FREESECT);
  const chainRegion = (start: number, count: number): void => {
    for (let k = 0; k < count; k++) fat[start + k] = k < count - 1 ? start + k + 1 : ENDOFCHAIN;
  };
  chainRegion(dirStart, dirSectors);
  if (miniFatSectors > 0) chainRegion(miniFatStart, miniFatSectors);
  if (miniStreamSectors > 0) chainRegion(miniStreamStart, miniStreamSectors);
  for (const big of bigStreams) chainRegion(big.entry.startSector, big.sectors);
  for (let k = 0; k < fatSectors; k++) fat[fatStart + k] = FATSECT;
  for (let k = 0; k < difatSectors; k++) fat[difatStart + k] = DIFSECT;

  // ── Serialize ──────────────────────────────────────────────────────────────────────────────────
  const buf = new Uint8Array((totalSectors + 1) * SECTOR); // +1 for the header sector
  const dv = new DataView(buf.buffer);
  const at = (sector: number): number => (sector + 1) * SECTOR;

  writeHeader(dv, {
    fatSectors,
    dirStart,
    miniFatStart,
    miniFatSectors,
    difatStart,
    difatSectors,
    fatStart,
  });

  for (const e of entries) writeDirEntry(dv, at(dirStart) + e.index * DIR_ENTRY_SIZE, e);

  for (let i = 0; i < miniFatSectors * FAT_ENTRIES_PER_SECTOR; i++) {
    dv.setUint32(at(miniFatStart) + i * 4, miniFat[i] ?? FREESECT, true);
  }
  // Guarded rather than looped-and-skipped: with no mini stream, miniStreamStart is ENDOFCHAIN and
  // `at()` of it is far outside the buffer, which a zero-length `set` would still reject.
  if (miniBytes.length > 0) buf.set(Uint8Array.from(miniBytes), at(miniStreamStart));
  for (const big of bigStreams) buf.set(big.data, at(big.entry.startSector));
  for (const [i, sector] of fat.entries()) dv.setUint32(at(fatStart) + i * 4, sector, true);

  // DIFAT sectors carry FAT-sector pointers 110.. (unreachable under the enforced bound, but the layout
  // is honoured: each DIFAT sector's tail points to the next, the last to ENDOFCHAIN).
  for (let d = 0; d < difatSectors; d++) {
    const base = at(difatStart + d);
    for (let i = 0; i < FAT_ENTRIES_PER_SECTOR - 1; i++) {
      const fatIdx = DIFAT_HEADER_SLOTS + d * (FAT_ENTRIES_PER_SECTOR - 1) + i;
      dv.setUint32(base + i * 4, fatIdx < fatSectors ? fatStart + fatIdx : FREESECT, true);
    }
    dv.setUint32(
      base + (FAT_ENTRIES_PER_SECTOR - 1) * 4,
      d < difatSectors - 1 ? difatStart + d + 1 : ENDOFCHAIN,
      true,
    );
  }

  return buf;
}

function validateSiblingNames(storageName: string, siblings: readonly CfbNode[]): void {
  const seen = new Set<string>();
  for (const node of siblings) {
    if (node.name.length === 0) throw new VbaAuthorError('entry name must not be empty');
    if (node.name.length > MAX_NAME_CHARS) {
      throw new VbaAuthorError(
        `entry name '${node.name}' exceeds the ${MAX_NAME_CHARS}-character CFB limit`,
      );
    }
    if (seen.has(node.name)) {
      throw new VbaAuthorError(
        `duplicate entry name '${node.name}' under storage '${storageName}'`,
      );
    }
    seen.add(node.name);
    if (!isStream(node)) validateSiblingNames(node.name, node.children);
  }
}

// Order a storage's children by the [MS-CFB] 2.6.4 comparison (shorter names sort first; ties broken by
// uppercased UTF-16 code units) and link them as a balanced binary tree. A host locates a child by
// walking this tree from the storage's `child` pointer, so the ordering and links must form a valid
// search tree.
//
// The recursion halves a slice rather than a lo/hi index pair, which makes "the slice is empty" and
// "there is no midpoint entry" the same observable fact: the `undefined` check *is* the base case.
function linkChildren(storage: DirEntry, kids: readonly DirEntry[]): void {
  const build = (nodes: readonly DirEntry[]): number => {
    const mid = nodes.length >> 1;
    const node = nodes[mid];
    if (node === undefined) return NOSTREAM;
    node.left = build(nodes.slice(0, mid));
    node.right = build(nodes.slice(mid + 1));
    return node.index;
  };
  storage.child = build([...kids].sort((a, b) => compareNames(a.name, b.name)));
}

function compareNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const ua = a.toUpperCase();
  const ub = b.toUpperCase();
  return ua < ub ? -1 : ua > ub ? 1 : 0;
}

function writeHeader(
  dv: DataView,
  p: {
    fatSectors: number;
    dirStart: number;
    miniFatStart: number;
    miniFatSectors: number;
    difatStart: number;
    difatSectors: number;
    fatStart: number;
  },
): void {
  dv.setUint32(0, 0xe011cfd0, true); // OLE2 signature (lo/hi)
  dv.setUint32(4, 0xe11ab1a1, true);
  dv.setUint16(24, 0x003e, true); // minor version
  dv.setUint16(26, 0x0003, true); // major version → v3 (512-byte sectors)
  dv.setUint16(28, 0xfffe, true); // little-endian byte order mark
  dv.setUint16(30, 9, true); // sector shift → 512
  dv.setUint16(32, 6, true); // mini-sector shift → 64
  dv.setUint32(44, p.fatSectors, true);
  dv.setUint32(48, p.dirStart, true);
  dv.setUint32(56, MINI_CUTOFF, true);
  dv.setUint32(60, p.miniFatStart, true);
  dv.setUint32(64, p.miniFatSectors, true);
  dv.setUint32(68, p.difatSectors > 0 ? p.difatStart : ENDOFCHAIN, true);
  dv.setUint32(72, p.difatSectors, true);
  // Header DIFAT: the first 109 FAT-sector pointers. Contiguous from fatStart under the enforced bound.
  for (let i = 0; i < DIFAT_HEADER_SLOTS; i++) {
    dv.setUint32(76 + i * 4, i < p.fatSectors ? p.fatStart + i : FREESECT, true);
  }
}

function writeDirEntry(dv: DataView, off: number, e: DirEntry): void {
  for (let i = 0; i < e.name.length; i++) dv.setUint16(off + i * 2, e.name.charCodeAt(i), true);
  dv.setUint16(off + 64, (e.name.length + 1) * 2, true); // name byte length incl. NUL terminator
  dv.setUint8(off + 66, e.type);
  dv.setUint8(off + 67, COLOR_BLACK);
  dv.setUint32(off + 68, e.left, true);
  dv.setUint32(off + 72, e.right, true);
  dv.setUint32(off + 76, e.child, true);
  dv.setUint32(off + 116, e.startSector, true);
  dv.setUint32(off + 120, e.size, true);
  // Size high (124), CLSID (80..95), state/time fields stay zero — valid for a v3 entry.
}
