import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {Workbook} from '../core/workbook.ts';
import {editXlsxVbaAddReference, editXlsxVbaRemoveModule} from '../io/xlsx/edit-vba.ts';
import {readXlsx} from '../io/xlsx/read.ts';
import {writeXlsx} from '../io/xlsx/write.ts';
import {CompoundFile} from './cfb.ts';
import {type CfbNode, writeCompoundFile} from './cfb-writer.ts';
import {VbaAuthorError, VbaParseError} from './errors.ts';
import {compressContainer, decompressContainer} from './ms-ovba.ts';
import {parseVbaProject} from './project.ts';
import {addVbaReference, removeVbaModule} from './project-editor.ts';

// ── Fixture builders ──────────────────────────────────────────────────────────────────────────────
// These construct a genuine, spec-valid `vbaProject.bin` from scratch: an MS-OVBA "store" encoder
// (literal-only chunks — valid compression that happens not to compress) and a minimal MS-CFB writer.
// No third-party bytes, no Excel. This exercises the real parse pipeline; the decoder is additionally
// pinned against an independent hand-verified vector (see the first decompress test) so the fixture and
// the code under test are not a closed loop on the copy-token path.

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;

/** Encode bytes as an MS-OVBA CompressedContainer using literal-only tokens (no back-references). */
function storeCompress(data: Uint8Array): Uint8Array {
  const out: number[] = [0x01];
  const CHUNK = 2048; // keep encoded chunk-data ≤ 4096 (the 12-bit size field) after 1/8 flag overhead
  for (let c = 0; c < data.length; c += CHUNK) {
    const slice = data.subarray(c, Math.min(c + CHUNK, data.length));
    const body: number[] = [];
    for (let i = 0; i < slice.length; i += 8) {
      body.push(0x00); // one flag byte, all-literal for the next up-to-8 bytes
      for (let j = i; j < Math.min(i + 8, slice.length); j++) body.push(slice[j] as number);
    }
    const header = 0xb000 | ((body.length - 1) & 0x0fff); // compressed + 0b011 sig + (size-1)
    out.push(header & 0xff, (header >> 8) & 0xff, ...body);
  }
  return Uint8Array.from(out);
}

function u16le(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff];
}
function u32le(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
}
function rec(id: number, data: number[]): number[] {
  return [...u16le(id), ...u32le(data.length), ...data];
}
function ascii(s: string): number[] {
  return [...s].map((c) => c.charCodeAt(0));
}
function utf16le(s: string): number[] {
  return [...s].flatMap((c) => u16le(c.charCodeAt(0)));
}

interface ModuleSpec {
  name: string;
  documentType: boolean; // MODULETYPE: false→procedural (0x21), true→non-procedural (0x22)
  sourceBytes: number[];
  pcodePrefixLen: number;
}

function buildDirStream(codePage: number, modules: ModuleSpec[]): number[] {
  const records: number[] = [];
  records.push(...rec(0x0003, u16le(codePage))); // PROJECTCODEPAGE
  // PROJECTVERSION: Size=4 counts only VersionMajor; the trailing 2-byte VersionMinor is uncounted —
  // the exact record that misaligns a naive TLV walk. Its presence proves the parser skips it.
  records.push(...rec(0x0009, u32le(0x04)), ...u16le(0x000a));
  records.push(...rec(0x000f, u16le(modules.length))); // MODULES_COUNT
  records.push(...rec(0x0013, u16le(0xffff))); // PROJECTCOOKIE
  for (const m of modules) {
    records.push(...rec(0x0019, ascii(m.name))); // MODULENAME
    records.push(...rec(0x001a, ascii(m.name))); // MODULESTREAMNAME (MBCS)
    records.push(...rec(0x0032, utf16le(m.name))); // Reserved: Unicode stream name — must be skipped
    records.push(...rec(0x0031, u32le(m.pcodePrefixLen))); // MODULEOFFSET
    records.push(...rec(m.documentType ? 0x0022 : 0x0021, [])); // MODULETYPE (Reserved u32 = Size 0)
    records.push(...rec(0x002b, [])); // MODULETERMINATOR
  }
  records.push(...rec(0x0010, [])); // dir Terminator — closes PROJECTMODULES, ends the dir stream
  return records;
}

/** Bytes stored in a module's CFB stream: a fake p-code prefix, then the compressed source. */
function buildModuleStream(m: ModuleSpec): Uint8Array {
  const compressed = storeCompress(Uint8Array.from(m.sourceBytes));
  const out = new Uint8Array(m.pcodePrefixLen + compressed.length);
  out.set(compressed, m.pcodePrefixLen); // prefix left as zeros — stand-in for the PerformanceCache
  return out;
}

/** A minimal MS-CFB container. Every stream here is < 4096 bytes, so all live in the mini stream. */
function buildCfb(streams: {name: string; data: Uint8Array}[]): Uint8Array {
  const SEC = 512;
  const MINI = 64;

  // Pack each stream into whole mini-sectors and chain them in the mini-FAT.
  const miniBytes: number[] = [];
  const miniFat: number[] = [];
  const entries = [{name: 'Root Entry', type: 5, start: 0, size: 0}];
  for (const s of streams) {
    assert.ok(s.data.length < 4096, 'fixture streams must be mini-stream sized');
    const startMini = miniBytes.length / MINI;
    const numMini = Math.max(1, Math.ceil(s.data.length / MINI));
    for (let k = 0; k < numMini; k++)
      miniFat.push(k < numMini - 1 ? startMini + k + 1 : ENDOFCHAIN);
    miniBytes.push(...s.data, ...new Array(numMini * MINI - s.data.length).fill(0));
    entries.push({name: s.name, type: 2, start: startMini, size: s.data.length});
  }

  const dirSectors = Math.max(1, Math.ceil((entries.length * 128) / SEC));
  const miniFatSectors = miniFat.length > 0 ? Math.ceil((miniFat.length * 4) / SEC) : 0;
  const miniStreamSectors = Math.ceil(miniBytes.length / SEC);

  let next = 0;
  const fatSectorIdx = next++;
  const dirStart = next;
  next += dirSectors;
  const miniFatStart = miniFatSectors > 0 ? next : ENDOFCHAIN;
  next += miniFatSectors;
  const miniStreamStart = next;
  next += miniStreamSectors;
  const totalSectors = next;
  assert.ok(totalSectors + 1 <= 128, 'fixture must fit a single FAT sector');

  const fat = new Array<number>(128).fill(FREESECT);
  fat[fatSectorIdx] = FATSECT;
  const chain = (from: number, count: number): void => {
    for (let k = 0; k < count; k++) fat[from + k] = k < count - 1 ? from + k + 1 : ENDOFCHAIN;
  };
  chain(dirStart, dirSectors);
  if (miniFatSectors > 0) chain(miniFatStart, miniFatSectors);
  chain(miniStreamStart, miniStreamSectors);

  entries[0]!.start = miniStreamStart;
  entries[0]!.size = miniBytes.length;

  const buf = new Uint8Array((totalSectors + 1) * SEC);
  const dv = new DataView(buf.buffer);
  const sectorOffset = (idx: number): number => (idx + 1) * SEC;

  // Header
  dv.setUint32(0, 0xe011cfd0, true);
  dv.setUint32(4, 0xe11ab1a1, true);
  dv.setUint16(24, 0x003e, true); // minor version
  dv.setUint16(26, 0x0003, true); // major version (v3)
  dv.setUint16(28, 0xfffe, true); // byte order
  dv.setUint16(30, 9, true); // sector shift → 512
  dv.setUint16(32, 6, true); // mini sector shift → 64
  dv.setUint32(44, 1, true); // number of FAT sectors
  dv.setUint32(48, dirStart, true); // first directory sector
  dv.setUint32(56, 4096, true); // mini-stream cutoff
  dv.setUint32(60, miniFatStart, true); // first mini-FAT sector
  dv.setUint32(64, miniFatSectors, true); // number of mini-FAT sectors
  dv.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector
  dv.setUint32(72, 0, true); // number of DIFAT sectors
  for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, i === 0 ? fatSectorIdx : FREESECT, true);

  // FAT sector
  for (let i = 0; i < 128; i++)
    dv.setUint32(sectorOffset(fatSectorIdx) + i * 4, fat[i] as number, true);

  // Directory
  entries.forEach((e, i) => {
    const off = sectorOffset(dirStart) + i * 128;
    const name16 = utf16le(e.name);
    name16.forEach((b, j) => {
      buf[off + j] = b;
    });
    dv.setUint16(off + 64, name16.length + 2, true); // name length incl. null terminator
    buf[off + 66] = e.type;
    dv.setUint32(off + 68, FREESECT, true); // left sibling  (tree unused by the reader)
    dv.setUint32(off + 72, FREESECT, true); // right sibling
    dv.setUint32(off + 76, FREESECT, true); // child
    dv.setUint32(off + 116, e.start, true);
    dv.setUint32(off + 120, e.size, true);
  });

  // Mini-FAT
  if (miniFatSectors > 0) {
    for (let i = 0; i < miniFatSectors * (SEC / 4); i++) {
      dv.setUint32(sectorOffset(miniFatStart) + i * 4, miniFat[i] ?? FREESECT, true);
    }
  }

  // Mini stream
  miniBytes.forEach((b, i) => {
    buf[sectorOffset(miniStreamStart) + i] = b;
  });

  return buf;
}

const PROJECT_STREAM = [
  'ID="{00000000-0000-0000-0000-000000000000}"',
  'Document=ThisWorkbook/&H00000000',
  'Module=Module1',
  'Class=Class1',
  '',
].join('\r\n');

function buildVbaProjectBin(codePage: number, modules: ModuleSpec[]): Uint8Array {
  const dir = storeCompress(Uint8Array.from(buildDirStream(codePage, modules)));
  return buildCfb([
    {name: 'PROJECT', data: strToU8(PROJECT_STREAM)},
    {name: 'dir', data: dir},
    ...modules.map((m) => ({name: m.name, data: buildModuleStream(m)})),
  ]);
}

// A three-module project: a document code-behind, a procedural .bas, and a class module — the last two
// share MODULETYPE 0x22/0x21 but are told apart by the PROJECT stream. Module1's source carries byte
// 0xC0, which is 'А' (U+0410) in code page 1251 — proving code-page-aware decoding, not latin1.
const MODULES: ModuleSpec[] = [
  {
    name: 'ThisWorkbook',
    documentType: true,
    sourceBytes: ascii('Private Sub Workbook_Open()\r\nEnd Sub'),
    pcodePrefixLen: 16,
  },
  {
    name: 'Module1',
    documentType: false,
    sourceBytes: [...ascii('Sub Test() '), 0xc0],
    pcodePrefixLen: 24,
  },
  {name: 'Class1', documentType: true, sourceBytes: ascii('Public X As Long'), pcodePrefixLen: 8},
];
const CODE_PAGE = 1251;

// ── Decompressor: independent ground-truth vector ───────────────────────────────────────────────────

test('decompressContainer expands a hand-verified copy-token vector', () => {
  // Container for "abcabc": literals a,b,c then CopyToken(offset 3, length 3). Encoded by hand from
  // [MS-OVBA] 2.4.1.3.19.3 — not by this suite's storeCompress, which never emits copy tokens.
  const container = Uint8Array.from([0x01, 0x05, 0xb0, 0x08, 0x61, 0x62, 0x63, 0x00, 0x20]);
  assert.equal(new TextDecoder('latin1').decode(decompressContainer(container)), 'abcabc');
});

test('decompressContainer round-trips the store encoder for multi-chunk data', () => {
  const data = new Uint8Array(5000).map((_, i) => (i * 31 + 7) & 0xff); // > 2048 → several chunks
  assert.deepEqual(decompressContainer(storeCompress(data)), data);
});

test('decompressContainer rejects a bad signature byte', () => {
  assert.throws(() => decompressContainer(Uint8Array.from([0x00, 0x01])), VbaParseError);
});

test('decompressContainer caps output to guard a decompression bomb', () => {
  const data = new Uint8Array(4096);
  assert.throws(() => decompressContainer(storeCompress(data), 0, 1024), VbaParseError);
});

// ── Compressor (§2.3b): compressContainer ────────────────────────────────────────────────────────────

test('compressContainer round-trips arbitrary data across the chunk boundary', () => {
  // Sizes straddling the 4096-byte chunk window and the 8-token flag group catch off-by-one framing.
  for (const n of [0, 1, 2, 3, 7, 8, 9, 100, 4095, 4096, 4097, 5000, 12000]) {
    const data = new Uint8Array(n).map((_, i) => (i * 131 + 7) & 0xff);
    assert.deepEqual(
      decompressContainer(compressContainer(data)),
      data,
      `round-trip failed at n=${n}`,
    );
  }
});

test('compressContainer emits copy tokens, shrinking repetitive data via run-length overlap', () => {
  const runs = new Uint8Array(4096).fill(0x41); // one byte repeated → a single overlapping back-reference
  const packed = compressContainer(runs);
  assert.ok(
    packed.length < 32,
    `4096 identical bytes should collapse to a tiny container, got ${packed.length}`,
  );
  assert.deepEqual(decompressContainer(packed), runs);

  const abab = Uint8Array.from({length: 6000}, (_, i) => (i % 2 ? 0x62 : 0x61));
  assert.ok(compressContainer(abab).length < abab.length / 4);
  assert.deepEqual(decompressContainer(compressContainer(abab)), abab);
});

test('compressContainer output re-parses as a real module through the whole pipeline', () => {
  // Compress genuine VBA source, wrap it as a module stream at offset 0, and read it back through the
  // production CFB writer + parser — the compressor feeding the reader end to end, no store-mode fixture.
  const source = 'Sub Demo()\r\n    MsgBox "hi"\r\n    MsgBox "hi"\r\nEnd Sub';
  const compressed = compressContainer(strToU8(source));
  const dir = compressContainer(
    Uint8Array.from(
      buildDirStream(1252, [
        {name: 'Demo', documentType: false, sourceBytes: [], pcodePrefixLen: 0},
      ]),
    ),
  );
  const bin = writeCompoundFile([
    {name: 'PROJECT', data: strToU8('Module=Demo\r\n')},
    {
      name: 'VBA',
      children: [
        {name: 'dir', data: dir},
        {name: 'Demo', data: compressed},
      ],
    },
  ]);
  const project = parseVbaProject(bin);
  assert.equal(
    project.modules[0]!.source,
    source,
    'the module source survives compress → write → parse',
  );
});

// ── CFB reader: malformed inputs fail closed ─────────────────────────────────────────────────────────

test('CompoundFile rejects a non-CFB signature', () => {
  assert.throws(() => new CompoundFile(new Uint8Array(512)), VbaParseError);
});

test('CompoundFile rejects a truncated header', () => {
  assert.throws(() => new CompoundFile(new Uint8Array(16)), VbaParseError);
});

test('CompoundFile rejects an out-of-range directory sector', () => {
  const bin = buildVbaProjectBin(CODE_PAGE, MODULES);
  const dv = new DataView(bin.buffer);
  dv.setUint32(48, 9999, true); // first directory sector points past the file
  assert.throws(() => new CompoundFile(bin), VbaParseError);
});

// ── Project parse ────────────────────────────────────────────────────────────────────────────────────

test('parseVbaProject decodes modules, code page, kinds, and source past the p-code', () => {
  const project = parseVbaProject(buildVbaProjectBin(CODE_PAGE, MODULES));

  assert.equal(project.codePage, 1251);
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
  assert.deepEqual(
    project.modules.map((m) => m.kind),
    ['document', 'procedural', 'class'], // Class1 (MODULETYPE 0x22) refined to 'class' via PROJECT
  );
  assert.match(project.modules[0]!.source, /Workbook_Open/);
  assert.ok(
    project.modules[1]!.source.includes('А'),
    'code page 1251 byte 0xC0 decodes to Cyrillic А',
  );
});

test('parseVbaProject throws VbaParseError on a corrupt dir stream', () => {
  const bin = buildVbaProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => parseVbaProject(bin.subarray(0, 900)), VbaParseError);
});

// ── CFB writer (§2.3a): writeCompoundFile ────────────────────────────────────────────────────────────

// Navigate the directory as a host does — from the Root Entry's child down each storage's balanced
// tree — collecting stream paths. Independent of CompoundFile, which linear-scans the directory and so
// would pass even over a broken tree; this asserts the tree Excel actually walks is a valid, acyclic
// search tree that reaches every entry.
function treeReachableStreams(bin: Uint8Array): string[] {
  const dv = new DataView(bin.buffer, bin.byteOffset, bin.byteLength);
  const at = (s: number): number => (s + 1) * 512;
  const dirStart = dv.getUint32(48, true);
  const NO = 0xffffffff;
  const entry = (i: number) => {
    const o = at(dirStart) + i * 128;
    const len = dv.getUint16(o + 64, true) / 2 - 1;
    let name = '';
    for (let k = 0; k < len; k++) name += String.fromCharCode(dv.getUint16(o + k * 2, true));
    return {
      name,
      type: bin[o + 66] as number,
      left: dv.getUint32(o + 68, true),
      right: dv.getUint32(o + 72, true),
      child: dv.getUint32(o + 76, true),
    };
  };
  const out: string[] = [];
  const seen = new Set<number>();
  const walk = (idx: number, prefix: string): void => {
    if (idx === NO) return;
    if (seen.has(idx)) throw new Error('cycle in directory tree');
    seen.add(idx);
    const e = entry(idx);
    walk(e.left, prefix);
    if (e.type === 2) out.push(`${prefix}/${e.name}`);
    if (e.type === 1) walk(e.child, `${prefix}/${e.name}`);
    walk(e.right, prefix);
  };
  walk(entry(0).child, '');
  return out;
}

test('writeCompoundFile round-trips mixed small, large, and empty streams through the reader', () => {
  const small = new Uint8Array(100).map((_, i) => i & 0xff);
  const large = new Uint8Array(9000).map((_, i) => (i * 7 + 3) & 0xff); // > 4096 cutoff → regular FAT
  const empty = new Uint8Array(0);
  const bin = writeCompoundFile([
    {name: 'small', data: small},
    {name: 'big', data: large},
    {name: 'empty', data: empty},
  ]);
  const cfb = new CompoundFile(bin);
  assert.deepEqual(cfb.readStream('small'), small);
  assert.deepEqual(
    cfb.readStream('big'),
    large,
    'a stream past the mini cutoff round-trips via the regular FAT',
  );
  assert.deepEqual(cfb.readStream('empty'), empty);
});

test('writeCompoundFile nests streams inside a storage and keeps the tree navigable', () => {
  const bin = writeCompoundFile([
    {name: 'PROJECT', data: strToU8('ID="x"')},
    {
      name: 'VBA',
      children: [
        {name: 'dir', data: Uint8Array.from([1, 2, 3])},
        {name: 'Module1', data: Uint8Array.from([4, 5, 6])},
      ],
    },
  ]);
  const cfb = new CompoundFile(bin);
  assert.deepEqual(cfb.readStream('dir'), Uint8Array.from([1, 2, 3]));
  assert.deepEqual(cfb.readStream('Module1'), Uint8Array.from([4, 5, 6]));

  assert.deepEqual(
    treeReachableStreams(bin).sort(),
    ['/PROJECT', '/VBA/Module1', '/VBA/dir'],
    'modules resolve under the VBA storage by tree navigation, not only by linear scan',
  );
});

test('writeCompoundFile produces a container parseVbaProject decodes', () => {
  // Build the VBA-project stream set from the same fixture bytes, but package it through the production
  // writer (proper VBA-storage hierarchy) rather than the test's buildCfb — proving the writer yields a
  // parseable project, not merely a reader-round-trippable blob.
  const dir = storeCompress(Uint8Array.from(buildDirStream(CODE_PAGE, MODULES)));
  const vbaChildren: CfbNode[] = [
    {name: 'dir', data: dir},
    ...MODULES.map((m) => ({name: m.name, data: buildModuleStream(m)})),
  ];
  const bin = writeCompoundFile([
    {name: 'PROJECT', data: strToU8(PROJECT_STREAM)},
    {name: 'VBA', children: vbaChildren},
  ]);

  const project = parseVbaProject(bin);
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
  assert.equal(project.codePage, 1251);
  assert.deepEqual(
    treeReachableStreams(bin)
      .filter((p) => p.startsWith('/VBA/'))
      .sort(),
    ['/VBA/Class1', '/VBA/Module1', '/VBA/ThisWorkbook', '/VBA/dir'],
  );
});

test('writeCompoundFile rejects an over-long stream name fail-closed', () => {
  assert.throws(
    () => writeCompoundFile([{name: 'x'.repeat(32), data: new Uint8Array(1)}]),
    VbaAuthorError,
  );
});

test('writeCompoundFile rejects duplicate sibling names', () => {
  assert.throws(
    () =>
      writeCompoundFile([
        {name: 'dup', data: new Uint8Array(1)},
        {name: 'dup', data: new Uint8Array(2)},
      ]),
    VbaAuthorError,
  );
});

// ── Workbook integration + round-trip non-regression ─────────────────────────────────────────────────

// One VBA-signature generation to wire into a package: the full relationship Type (carrying the year
// segment the generation really uses — 2006 legacy, 2014 agile, 2020 V3), the sibling part's file name
// and content type, and its raw bytes. Office attaches up to three of these over one project.
interface SignatureGeneration {
  readonly relType: string;
  readonly fileName: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
}

const legacySignature = (bytes: Uint8Array): SignatureGeneration => ({
  relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature',
  fileName: 'vbaProjectSignature.bin',
  contentType: 'application/vnd.ms-office.vbaProjectSignature',
  bytes,
});
const agileSignature = (bytes: Uint8Array): SignatureGeneration => ({
  relType: 'http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile',
  fileName: 'vbaProjectSignatureAgile.bin',
  contentType: 'application/vnd.ms-office.vbaProjectSignatureAgile',
  bytes,
});
const v3Signature = (bytes: Uint8Array): SignatureGeneration => ({
  relType: 'http://schemas.microsoft.com/office/2020/relationships/vbaProjectSignatureV3',
  fileName: 'vbaProjectSignatureV3.bin',
  contentType: 'application/vnd.ms-office.vbaProjectSignatureV3',
  bytes,
});

// A minimal macro-enabled package around a vbaProject.bin. Pass one or more `sigs` to additionally wire
// digital signatures over the project — sibling parts reached from the project part's own rels, the
// shape an authoring replace must invalidate (a signature over old bytes cannot vouch for new ones).
function xlsmPackage(vbaBin: Uint8Array, sigs: readonly SignatureGeneration[] = []): Uint8Array {
  const ms = 'http://schemas.microsoft.com/office/2006/relationships';
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>' +
        sigs
          .map((s) => `<Override PartName="/xl/${s.fileName}" ContentType="${s.contentType}"/>`)
          .join('') +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>',
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        `<Relationship Id="rId2" Type="${ms}/vbaProject" Target="vbaProject.bin"/>` +
        '</Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    ),
    'xl/vbaProject.bin': vbaBin,
  };
  if (sigs.length > 0) {
    files['xl/_rels/vbaProject.bin.rels'] = strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        sigs
          .map(
            (s, i) => `<Relationship Id="rId${i + 1}" Type="${s.relType}" Target="${s.fileName}"/>`,
          )
          .join('') +
        '</Relationships>',
    );
    for (const s of sigs) files[`xl/${s.fileName}`] = s.bytes;
  }
  return zipSync(files);
}

test('Workbook.vbaProject decodes macros from a read .xlsm and memoises', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  const project = wb.vbaProject;
  assert.ok(project, 'a macro-enabled workbook exposes its vbaProject');
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
  assert.equal(wb.vbaProject, project, 'the parsed project is memoised, not re-decoded');
});

test('Workbook.vbaProject is undefined for a macro-free workbook', () => {
  const plain = zipSync(unzipSync(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES))));
  const files = unzipSync(plain);
  delete files['xl/vbaProject.bin'];
  // Drop the vbaProject relationship too so the package is a consistent macro-free book.
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  );
  assert.equal(readXlsx(zipSync(files)).vbaProject, undefined);
});

test('reading vbaProject does not regress byte-for-byte macro preservation on write', () => {
  const vbaBin = buildVbaProjectBin(CODE_PAGE, MODULES);
  const wb = readXlsx(xlsmPackage(vbaBin));
  // Force the read-only projection before writing — it must not perturb the preserved bytes.
  assert.ok(wb.vbaProject);
  const out = unzipSync(writeXlsx(wb));
  const reBin = Object.entries(out).find(([n]) => /vbaProject\.bin$/.test(n))?.[1];
  assert.ok(reBin, 'the written package still carries a vbaProject.bin');
  assert.deepEqual(reBin, vbaBin, 'the macro blob is re-emitted byte-for-byte');
  // And it still parses from the re-emitted package.
  assert.deepEqual(
    readXlsx(writeXlsx(wb)).vbaProject?.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
});

// ── Attach-blob authoring (§2.1): Workbook.vbaProjectBytes get/set ───────────────────────────────────

test('attaching vbaProjectBytes turns a plain workbook macro-enabled and embeds the blob verbatim', () => {
  const bin = buildVbaProjectBin(CODE_PAGE, MODULES);
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  wb.vbaProjectBytes = bin;

  const out = unzipSync(writeXlsx(wb));
  assert.ok(out['xl/vbaProject.bin'], 'the written package carries a vbaProject.bin');
  assert.deepEqual(out['xl/vbaProject.bin'], bin, 'the attached blob is embedded byte-for-byte');

  const ct = new TextDecoder().decode(out['[Content_Types].xml']);
  assert.match(
    ct,
    /application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml/,
    'the workbook part is declared macro-enabled',
  );
  assert.match(ct, /application\/vnd\.ms-office\.vbaProject/, 'the .bin is typed as a vbaProject');

  // The re-read package exposes the same macros.
  assert.deepEqual(
    readXlsx(writeXlsx(wb)).vbaProject?.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
});

test('vbaProjectBytes copies a macro project from one workbook to another', () => {
  const source = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  const target = new Workbook();
  target.addWorksheet('Sheet1');

  const bytes = source.vbaProjectBytes;
  assert.ok(bytes, 'the source workbook exposes its raw macro blob');
  target.vbaProjectBytes = bytes;

  assert.deepEqual(
    readXlsx(writeXlsx(target)).vbaProject?.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'the copied project decodes from the target package',
  );
});

test('the vbaProjectBytes getter returns a defensive copy', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  const first = wb.vbaProjectBytes;
  assert.ok(first);
  first.fill(0); // scribble on the returned copy
  const second = wb.vbaProjectBytes;
  assert.ok(second);
  assert.notDeepEqual(second, first, 'mutating a returned copy does not corrupt the stored blob');
  // The stored blob still round-trips and parses.
  assert.ok(readXlsx(writeXlsx(wb)).vbaProject);
});

test('assigning undefined removes the macro project, reverting to a plain package', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  assert.ok(wb.vbaProject, 'precondition: the workbook has macros');
  wb.vbaProjectBytes = undefined;

  assert.equal(wb.vbaProject, undefined, 'the read view reflects the removal');
  assert.equal(wb.vbaProjectBytes, undefined, 'no blob remains attached');

  const out = unzipSync(writeXlsx(wb));
  assert.equal(
    out['xl/vbaProject.bin'],
    undefined,
    'the package no longer carries a vbaProject.bin',
  );
  const ct = new TextDecoder().decode(out['[Content_Types].xml']);
  assert.doesNotMatch(ct, /macroEnabled/, 'the workbook is no longer declared macro-enabled');
});

test('attaching a malformed blob is rejected fail-closed and leaves the workbook untouched', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  const original = wb.vbaProjectBytes;

  assert.throws(() => {
    wb.vbaProjectBytes = Uint8Array.from([1, 2, 3, 4]); // not a CFB container
  }, VbaParseError);

  // The reject happens before the old project is cleared, so the workbook is unchanged.
  assert.deepEqual(
    wb.vbaProjectBytes,
    original,
    'a rejected attach does not disturb the existing blob',
  );
});

test('replacing the project drops a now-stale signature over the old bytes', () => {
  const oldBin = buildVbaProjectBin(CODE_PAGE, MODULES);
  const sig = Uint8Array.from([0xde, 0xad, 0xbe, 0xef]);
  const wb = readXlsx(xlsmPackage(oldBin, [legacySignature(sig)]));
  // Precondition: the signature part is present in the read package.
  assert.ok(
    unzipSync(writeXlsx(wb))['xl/vbaProjectSignature.bin'],
    'signature present before replace',
  );
  assert.equal(wb.vbaProjectSigned, true, 'the accessor sees the signature before replace');

  const newBin = buildVbaProjectBin(1252, [
    {
      name: 'Module1',
      documentType: false,
      sourceBytes: ascii('Sub Fresh()\r\nEnd Sub'),
      pcodePrefixLen: 8,
    },
  ]);
  wb.vbaProjectBytes = newBin;

  const out = unzipSync(writeXlsx(wb));
  assert.deepEqual(out['xl/vbaProject.bin'], newBin, 'the new blob is embedded');
  assert.equal(
    out['xl/vbaProjectSignature.bin'],
    undefined,
    'the stale signature over the old bytes is dropped, not left to advertise a broken signature',
  );
  assert.equal(
    wb.vbaProjectSigned,
    false,
    'the drop is visible through the accessor: a replaced project reads unsigned',
  );
  assert.deepEqual(wb.vbaProjectSignatures, [], 'no signatures remain after replace');
  assert.deepEqual(
    readXlsx(writeXlsx(wb)).vbaProject?.modules.map((m) => m.name),
    ['Module1'],
    'the replacement project decodes',
  );
});

test('Workbook.vbaProjectSigned reports a signed project and exposes the raw signature bytes', () => {
  const sig = Uint8Array.from([0xde, 0xad, 0xbe, 0xef, 4, 5]);
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES), [legacySignature(sig)]));

  assert.equal(wb.vbaProjectSigned, true);
  assert.equal(wb.vbaProjectSignatures.length, 1);
  const only = wb.vbaProjectSignatures[0];
  assert.ok(only);
  assert.equal(only.kind, 'legacy');
  assert.deepEqual(only.bytes, sig, 'the raw signature blob is passed through verbatim');

  // The getter hands back a defensive copy; mutating it must not corrupt the preserved bytes.
  only.bytes[0] = 0;
  assert.equal(
    wb.vbaProjectSignatures[0]?.bytes[0],
    0xde,
    'the stored signature bytes are unchanged',
  );
});

test('Workbook.vbaProjectSigned reads false for an unsigned macro project', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  assert.equal(wb.vbaProjectSigned, false);
  assert.deepEqual(wb.vbaProjectSignatures, []);
});

test('Workbook.vbaProjectSigned reads false for a workbook with no macros', () => {
  const wb = new Workbook();
  assert.equal(wb.vbaProjectSigned, false);
  assert.deepEqual(wb.vbaProjectSignatures, []);
});

test('Workbook.vbaProjectSignatures reports every generation, across the years each URI carries', () => {
  const wb = readXlsx(
    xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES), [
      legacySignature(Uint8Array.from([1])),
      agileSignature(Uint8Array.from([2])),
      v3Signature(Uint8Array.from([3])),
    ]),
  );

  assert.equal(wb.vbaProjectSigned, true);
  // Detection keys off the relationship Type's final segment, so all three are found even though their
  // URIs carry different year segments (2006 / 2014 / 2020); the closure walk preserves each part.
  assert.deepEqual(
    wb.vbaProjectSignatures.map((s) => s.kind),
    ['legacy', 'agile', 'v3'],
  );
  assert.deepEqual(
    wb.vbaProjectSignatures.map((s) => [...s.bytes]),
    [[1], [2], [3]],
  );
});

// ── Shared fixtures for structural edits: removeVbaModule / addVbaReference ────────────────────────────

// The PROJECTwm stream pairs each module's MBCS name with its UTF-16 name, both NUL-terminated, ending
// with an empty pair — one record per module the dir/PROJECT streams declare, so removeVbaModule's
// splice (which counts existing records against the parsed module count) has a real structure to shrink.
function buildProjectwmStream(names: readonly string[]): Uint8Array {
  const b: number[] = [];
  for (const name of names) b.push(...ascii(name), 0x00, ...utf16le(name), 0x00, 0x00);
  b.push(0x00, 0x00);
  return Uint8Array.from(b);
}

// Package a project through the *production* CFB writer so it has the navigable red-black sibling tree
// the removeVbaModule/addVbaReference editors walk (buildVbaProjectBin leaves those links null — fine for
// the linear-scan reader, but the editors rebuild the tree). Optionally append raw dir records (e.g. a
// PROJECTREFERENCES entry) and a distinctive _VBA_PROJECT, so a test can prove both survive / are
// replaced as intended.
function buildNavigableProjectBin(
  codePage: number,
  modules: ModuleSpec[],
  extraDirRecords: number[] = [],
): Uint8Array {
  const dir = storeCompress(
    Uint8Array.from([...buildDirStream(codePage, modules), ...extraDirRecords]),
  );
  const vbaChildren: CfbNode[] = [
    {name: 'dir', data: dir},
    {name: '_VBA_PROJECT', data: Uint8Array.from([0x61, 0xcc, 0x5e, 0x00, 0x00, 0x01, 0x02, 0x03])},
    ...modules.map((m) => ({name: m.name, data: buildModuleStream(m)})),
  ];
  return writeCompoundFile([
    {name: 'PROJECT', data: strToU8(PROJECT_STREAM)},
    {name: 'PROJECTwm', data: buildProjectwmStream(modules.map((m) => m.name))},
    {name: 'VBA', children: vbaChildren},
  ]);
}

function indexOfBytes(haystack: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = 0; i + needle.length <= haystack.length; i++) {
    for (let j = 0; j < needle.length; j++) if (haystack[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}

// The dir stream's MODULES_COUNT field ([MS-OVBA] 2.3.4.2.3.2) — not cross-checked by parseVbaProject
// (which discovers modules by MODULETERMINATOR markers regardless of the count), but Excel relies on it,
// so removeVbaModule must keep it in sync. Reads it directly out of the decompressed dir bytes.
function readModulesCount(dir: Uint8Array): number {
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const id = (dir[pos] as number) | ((dir[pos + 1] as number) << 8);
    const size =
      ((dir[pos + 2] as number) |
        ((dir[pos + 3] as number) << 8) |
        ((dir[pos + 4] as number) << 16) |
        ((dir[pos + 5] as number) << 24)) >>>
      0;
    const dataStart = pos + 6;
    if (id === 0x000f) return (dir[dataStart] as number) | ((dir[dataStart + 1] as number) << 8);
    pos = dataStart + size;
    if (id === 0x0009) pos += 2; // PROJECTVERSION's uncounted VersionMinor
  }
  throw new Error('MODULES_COUNT not found');
}

// ── Structural edit: removeVbaModule ────────────────────────────────────────────────────────────────

test('removeVbaModule removes a procedural module, preserving references and untouched modules', () => {
  const refPayload = ascii('*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42');
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES, rec(0x000d, refPayload));

  const removed = removeVbaModule(bin, 'Module1');

  const project = parseVbaProject(removed);
  assert.deepEqual(
    project.modules.map((m) => [m.name, m.kind]),
    [
      ['ThisWorkbook', 'document'],
      ['Class1', 'class'],
    ],
    'Module1 is gone; the other modules and their order survive',
  );

  const before = new CompoundFile(bin);
  const after = new CompoundFile(removed);
  for (const name of ['ThisWorkbook', 'Class1']) {
    assert.deepEqual(
      after.readStream(name),
      before.readStream(name),
      `${name} rides through unchanged`,
    );
  }
  assert.equal(after.readStream('Module1'), undefined, "Module1's stream is gone");

  // _VBA_PROJECT is preserved untouched — Excel runs the modules' existing p-code, and resetting the
  // cookie would crash the load.
  assert.deepEqual(after.readStream('_VBA_PROJECT'), before.readStream('_VBA_PROJECT'));

  const dirBefore = decompressContainer(before.readStream('dir')!);
  const dirAfter = decompressContainer(after.readStream('dir')!);
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(refPayload)) >= 0,
    'the PROJECTREFERENCES record is preserved',
  );
  assert.equal(readModulesCount(dirBefore), 3);
  assert.equal(
    readModulesCount(dirAfter),
    2,
    'MODULES_COUNT is decremented for the removed module',
  );

  const projectText = strFromU8(after.readStream('PROJECT')!);
  assert.doesNotMatch(projectText, /^Module=Module1$/m, "Module1's declaration line is gone");
  assert.match(projectText, /^Document=ThisWorkbook\/&H00000000$/m, 'ThisWorkbook line survives');
  assert.match(projectText, /^Class=Class1$/m, 'Class1 line survives');

  assert.equal(
    indexOfBytes(after.readStream('PROJECTwm')!, Uint8Array.from(ascii('Module1'))),
    -1,
    'PROJECTwm no longer carries the removed module name',
  );

  assert.deepEqual(
    treeReachableStreams(removed).sort(),
    ['/PROJECT', '/PROJECTwm', '/VBA/Class1', '/VBA/ThisWorkbook', '/VBA/_VBA_PROJECT', '/VBA/dir'],
    'the removed module no longer resolves under the VBA storage by tree navigation',
  );
});

test('removeVbaModule matches the module name case-insensitively', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  const removed = removeVbaModule(bin, 'module1');
  assert.deepEqual(
    parseVbaProject(removed).modules.map((m) => m.name),
    ['ThisWorkbook', 'Class1'],
  );
});

test('removeVbaModule rejects an unknown module name', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => removeVbaModule(bin, 'NoSuchModule'), VbaAuthorError);
});

test('removeVbaModule rejects removing a document module fail-closed', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => removeVbaModule(bin, 'ThisWorkbook'), VbaAuthorError);
});

test('removeVbaModule rejects a malformed container as a parse error', () => {
  assert.throws(() => removeVbaModule(Uint8Array.from([1, 2, 3, 4]), 'Module1'), VbaParseError);
});

// ── Structural edit: addVbaReference ─────────────────────────────────────────────────────────────────

// Microsoft Scripting Runtime's real GUID/path — the exact reference this splice was verified against on
// a genuine Excel-authored project (2026-07-23, excel-gui-automation probe, ADR 0012/0013 provenance).
const SCRIPTING_REF = {
  name: 'Scripting',
  displayName: 'Microsoft Scripting Runtime',
  guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
  majorVersion: 1,
  minorVersion: 0,
  path: 'C:\\Windows\\System32\\scrrun.dll',
};
const SCRIPTING_LIBID =
  '*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime';

test('addVbaReference adds a registered reference to a project with no existing reference, without touching PROJECT/PROJECTwm', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  const added = addVbaReference(bin, SCRIPTING_REF);

  // Modules are unaffected — same set, same order.
  assert.deepEqual(
    parseVbaProject(added).modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );

  const before = new CompoundFile(bin);
  const after = new CompoundFile(added);
  const dirAfter = decompressContainer(after.readStream('dir')!);
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(ascii(SCRIPTING_LIBID))) >= 0,
    'the assembled Libid string is present in the dir stream',
  );
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(ascii('Scripting'))) >= 0,
    'the REFERENCENAME name is present',
  );
  assert.equal(readModulesCount(dirAfter), 3, 'MODULES_COUNT is untouched by adding a reference');

  // No real Excel-authored PROJECT stream carries a Reference= line for a registered library reference
  // (verified against a genuine Excel-authored project) — so neither PROJECT nor PROJECTwm changes here.
  assert.deepEqual(after.readStream('PROJECT'), before.readStream('PROJECT'));
  assert.deepEqual(after.readStream('PROJECTwm'), before.readStream('PROJECTwm'));

  // _VBA_PROJECT is preserved untouched — Excel runs the modules' existing p-code, and resetting the
  // cookie would crash the load.
  assert.deepEqual(after.readStream('_VBA_PROJECT'), before.readStream('_VBA_PROJECT'));
});

test('addVbaReference adds a reference to an existing project, preserving an existing reference and every module byte-for-byte', () => {
  const existingRefPayload = ascii('*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42');
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES, rec(0x000d, existingRefPayload));

  const added = addVbaReference(bin, SCRIPTING_REF);

  const before = new CompoundFile(bin);
  const after = new CompoundFile(added);
  for (const name of ['ThisWorkbook', 'Module1', 'Class1']) {
    assert.deepEqual(
      after.readStream(name),
      before.readStream(name),
      `${name} rides through unchanged`,
    );
  }

  const dirBefore = decompressContainer(before.readStream('dir')!);
  const dirAfter = decompressContainer(after.readStream('dir')!);
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(existingRefPayload)) >= 0,
    'the pre-existing reference is preserved',
  );
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(ascii(SCRIPTING_LIBID))) >= 0,
    'the new reference is present',
  );
  assert.equal(readModulesCount(dirBefore), 3);
  assert.equal(readModulesCount(dirAfter), 3, 'MODULES_COUNT is unaffected by adding a reference');

  assert.deepEqual(
    parseVbaProject(added).modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'the module set and order are unaffected',
  );
});

test('addVbaReference rejects an invalid reference name', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, name: '1Bad'}), VbaAuthorError);
});

test('addVbaReference rejects a malformed GUID', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, guid: 'not-a-guid'}), VbaAuthorError);
});

test('addVbaReference rejects an out-of-range version or LCID', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, majorVersion: -1}), VbaAuthorError);
  assert.throws(
    () => addVbaReference(bin, {...SCRIPTING_REF, minorVersion: 0x10000}),
    VbaAuthorError,
  );
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, lcid: -1}), VbaAuthorError);
});

test('addVbaReference rejects an invalid path', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, path: ''}), VbaAuthorError);
  assert.throws(
    () => addVbaReference(bin, {...SCRIPTING_REF, path: 'C:\\has#hash.dll'}),
    VbaAuthorError,
  );
});

test('addVbaReference rejects display text the code page cannot represent', () => {
  const bin = buildNavigableProjectBin(1252, MODULES);
  assert.throws(
    () => addVbaReference(bin, {...SCRIPTING_REF, displayName: '你好'}),
    VbaAuthorError,
  );
});

test('addVbaReference rejects a malformed container as a parse error', () => {
  assert.throws(() => addVbaReference(Uint8Array.from([1, 2, 3, 4]), SCRIPTING_REF), VbaParseError);
});

// ── Structural edits through the public surface: Workbook.removeVbaModule / addVbaReference ───────────

test('Workbook.removeVbaModule removes a module from a read workbook and preserves references end-to-end', () => {
  const refPayload = ascii('*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42');
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES, rec(0x000d, refPayload));
  const wb = readXlsx(xlsmPackage(bin));

  wb.removeVbaModule('Module1');

  const reDir = decompressContainer(
    new CompoundFile(unzipSync(writeXlsx(wb))['xl/vbaProject.bin']!).readStream('dir')!,
  );
  assert.ok(
    indexOfBytes(reDir, Uint8Array.from(refPayload)) >= 0,
    'the PROJECTREFERENCES record survives the whole package round-trip',
  );

  const reread = readXlsx(writeXlsx(wb));
  assert.deepEqual(
    reread.vbaProject?.modules.map((m) => [m.name, m.kind]),
    [
      ['ThisWorkbook', 'document'],
      ['Class1', 'class'],
    ],
  );
});

test('Workbook.removeVbaModule on a macro-free workbook throws', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(() => wb.removeVbaModule('Module1'), VbaAuthorError);
});

test('Workbook.removeVbaModule rejects an unknown module without disturbing the existing project', () => {
  const wb = readXlsx(xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES)));
  const before = wb.vbaProjectBytes;

  assert.throws(() => wb.removeVbaModule('Nope'), VbaAuthorError);
  assert.deepEqual(wb.vbaProjectBytes, before, 'a rejected removal leaves the workbook untouched');
});

test('Workbook.removeVbaModule rejects removing a document module fail-closed', () => {
  const wb = readXlsx(xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES)));
  assert.throws(() => wb.removeVbaModule('ThisWorkbook'), VbaAuthorError);
});

test('Workbook.addVbaReference adds a reference to a read workbook, preserving modules', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  const wb = readXlsx(xlsmPackage(bin));

  wb.addVbaReference({
    name: 'Scripting',
    guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
    majorVersion: 1,
    minorVersion: 0,
    path: 'C:\\Windows\\System32\\scrrun.dll',
  });

  const reDir = decompressContainer(
    new CompoundFile(unzipSync(writeXlsx(wb))['xl/vbaProject.bin']!).readStream('dir')!,
  );
  assert.ok(
    indexOfBytes(reDir, Uint8Array.from(ascii('Scripting'))) >= 0,
    'the new reference is present after a full package round-trip',
  );
  assert.deepEqual(
    readXlsx(writeXlsx(wb)).vbaProject?.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'the module set is unaffected',
  );
});

test('Workbook.addVbaReference on a macro-free workbook throws', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(
    () =>
      wb.addVbaReference({
        name: 'Scripting',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
    VbaAuthorError,
  );
});

test('Workbook.addVbaReference rejects an invalid reference without disturbing the existing project', () => {
  const wb = readXlsx(xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES)));
  const before = wb.vbaProjectBytes;

  assert.throws(
    () =>
      wb.addVbaReference({
        name: '1Bad',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
    VbaAuthorError,
  );
  assert.deepEqual(wb.vbaProjectBytes, before, 'a rejected add leaves the workbook untouched');
});

// ── Package-preserving structural edits: editXlsxVbaRemoveModule / editXlsxVbaAddReference ─────────────

test('editXlsxVbaRemoveModule removes a module and preserves every other package part byte-for-byte', () => {
  const refPayload = ascii('*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42');
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES, rec(0x000d, refPayload)));
  const before = unzipSync(pkg);

  const after = unzipSync(editXlsxVbaRemoveModule(pkg, 'Module1'));

  assert.deepEqual(
    Object.keys(after).sort(),
    Object.keys(before).sort(),
    'no parts are added or removed',
  );
  for (const name of Object.keys(before)) {
    if (name === 'xl/vbaProject.bin') continue;
    assert.deepEqual(after[name], before[name], `${name} is preserved byte-for-byte`);
  }

  const project = parseVbaProject(after['xl/vbaProject.bin']!);
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Class1'],
  );
  const reDir = decompressContainer(
    new CompoundFile(after['xl/vbaProject.bin']!).readStream('dir')!,
  );
  assert.ok(
    indexOfBytes(reDir, Uint8Array.from(refPayload)) >= 0,
    'the PROJECTREFERENCES record survives the splice',
  );
});

test('editXlsxVbaRemoveModule drops a stale signature part', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES), [
    legacySignature(Uint8Array.from([1, 2, 3])),
  ]);
  assert.ok(unzipSync(pkg)['xl/vbaProjectSignature.bin'], 'precondition: the package is signed');

  const after = unzipSync(editXlsxVbaRemoveModule(pkg, 'Module1'));
  assert.equal(after['xl/vbaProjectSignature.bin'], undefined, 'the signature part is dropped');
});

test('editXlsxVbaRemoveModule throws for a macro-free package', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(() => editXlsxVbaRemoveModule(writeXlsx(wb), 'Module1'), VbaAuthorError);
});

test('editXlsxVbaRemoveModule propagates VbaAuthorError for an unknown module', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));
  assert.throws(() => editXlsxVbaRemoveModule(pkg, 'Nope'), VbaAuthorError);
});

test('editXlsxVbaRemoveModule propagates VbaAuthorError for a document module', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));
  assert.throws(() => editXlsxVbaRemoveModule(pkg, 'ThisWorkbook'), VbaAuthorError);
});

test('editXlsxVbaAddReference adds a reference and preserves every other package part byte-for-byte', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));
  const before = unzipSync(pkg);

  const after = unzipSync(
    editXlsxVbaAddReference(pkg, {
      name: 'Scripting',
      guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
      majorVersion: 1,
      minorVersion: 0,
      path: 'C:\\Windows\\System32\\scrrun.dll',
    }),
  );

  for (const name of Object.keys(before)) {
    if (name === 'xl/vbaProject.bin') continue;
    assert.deepEqual(after[name], before[name], `${name} is preserved byte-for-byte`);
  }

  const reDir = decompressContainer(
    new CompoundFile(after['xl/vbaProject.bin']!).readStream('dir')!,
  );
  assert.ok(
    indexOfBytes(reDir, Uint8Array.from(ascii('Scripting'))) >= 0,
    'the new reference is present in the spliced project',
  );
  assert.deepEqual(
    parseVbaProject(after['xl/vbaProject.bin']!).modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'the module set is unaffected',
  );
});

test('editXlsxVbaAddReference throws for a macro-free package', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(
    () =>
      editXlsxVbaAddReference(writeXlsx(wb), {
        name: 'Scripting',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
    VbaAuthorError,
  );
});

test('editXlsxVbaAddReference propagates VbaAuthorError for an invalid reference', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));
  assert.throws(
    () =>
      editXlsxVbaAddReference(pkg, {
        name: '1Bad',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
    VbaAuthorError,
  );
});
