import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {strToU8, unzipSync, zipSync} from 'fflate';

import {readXlsx} from '../io/xlsx/read.ts';
import {writeXlsx} from '../io/xlsx/write.ts';
import {CompoundFile} from './cfb.ts';
import {VbaParseError} from './errors.ts';
import {decompressContainer} from './ms-ovba.ts';
import {parseVbaProject} from './project.ts';

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
  for (const m of modules) {
    records.push(...rec(0x0019, ascii(m.name))); // MODULENAME
    records.push(...rec(0x001a, ascii(m.name))); // MODULESTREAMNAME (MBCS)
    records.push(...rec(0x0032, utf16le(m.name))); // Reserved: Unicode stream name — must be skipped
    records.push(...rec(0x0031, u32le(m.pcodePrefixLen))); // MODULEOFFSET
    records.push(...rec(m.documentType ? 0x0022 : 0x0021, [])); // MODULETYPE (Reserved u32 = Size 0)
    records.push(...rec(0x002b, [])); // MODULETERMINATOR
  }
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

// ── Workbook integration + round-trip non-regression ─────────────────────────────────────────────────

function xlsmPackage(vbaBin: Uint8Array): Uint8Array {
  const rel = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>' +
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
        `<Relationship Id="rId2" Type="${rel}" Target="vbaProject.bin"/>` +
        '</Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    ),
    'xl/vbaProject.bin': vbaBin,
  });
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
