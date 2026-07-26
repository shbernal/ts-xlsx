// A hand-built VBA project fixture.
//
// The .xlsm cases need a *structurally real* vbaProject.bin — a PROJ record and a
// `document` code-behind module — because those are exactly the parts splice-editing must
// preserve and re-synthesis cannot. Building it here keeps the bytes auditable.

import {strToU8, zipSync} from 'fflate';
import {compressContainer, writeCompoundFile} from './runtime.ts';

// ── Macro-enabled fixture builder ──────────────────────────────────────────────────────────────────
// Assemble a genuine, navigable vbaProject.bin (via the production CFB writer + MS-OVBA compressor, so
// its storage tree is walkable exactly as Excel's is) inside a minimal .xlsm package. This is the only
// way to produce an edit-in-place *input* without an interactive VBA editor: the writer cannot author a
// project from a model (no reference support, document-module linkage is host-coupled), but the editor
// splices new module source into an existing bin. The fixture carries a hand-crafted PROJECTREFERENCES
// record and a `document` code-behind module — the two things splice-editing must preserve that
// re-synthesis structurally cannot.
export const vbaU16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
export const vbaU32 = (n: number) => [
  n & 0xff,
  (n >> 8) & 0xff,
  (n >> 16) & 0xff,
  (n >> 24) & 0xff,
];
export const vbaAscii = (s: string) => [...s].map((c) => c.charCodeAt(0));
export const vbaUtf16 = (s: string) => [...s].flatMap((c) => vbaU16(c.charCodeAt(0)));
export const vbaRec = (id: number, data: number[]) => [
  ...vbaU16(id),
  ...vbaU32(data.length),
  ...data,
];

export interface VbaFixtureModule {
  readonly name: string;
  readonly document: boolean;
  readonly sourceBytes: number[];
  readonly pcodePrefixLen: number;
}

// Module1's source carries byte 0xC0 — 'А' (U+0410) in code page 1251 — so a code-page-blind (latin1)
// reader would corrupt it; preserving it byte-for-byte proves the untouched stream rides through raw.
export const VBA_FIXTURE_MODULES: VbaFixtureModule[] = [
  {
    name: 'ThisWorkbook',
    document: true,
    sourceBytes: vbaAscii('Private Sub Workbook_Open()\r\nEnd Sub'),
    pcodePrefixLen: 16,
  },
  {
    name: 'Module1',
    document: false,
    sourceBytes: [...vbaAscii('Sub Test() '), 0xc0],
    pcodePrefixLen: 24,
  },
  {name: 'Class1', document: true, sourceBytes: vbaAscii('Public X As Long'), pcodePrefixLen: 8},
];
export const VBA_FIXTURE_CODE_PAGE = 1251;
// A REFERENCEREGISTERED record ([MS-OVBA] 2.3.4.2). Its distinctive marker must survive the edit
// verbatim — the writer emits no references at all, so its presence proves splice-not-resynthesize.
export const VBA_FIXTURE_REF_MARKER = '*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42';

export const VBA_PROJECT_STREAM = [
  'ID="{00000000-0000-0000-0000-000000000000}"',
  'Document=ThisWorkbook/&H00000000',
  'Module=Module1',
  'Class=Class1',
  '',
].join('\r\n');

export function buildVbaFixtureBin(): Uint8Array {
  const dir: number[] = [];
  dir.push(...vbaRec(0x0003, vbaU16(VBA_FIXTURE_CODE_PAGE))); // PROJECTCODEPAGE
  dir.push(...vbaRec(0x0009, vbaU32(0x04)), ...vbaU16(0x000a)); // PROJECTVERSION (uncounted minor)
  dir.push(...vbaRec(0x000f, vbaU16(VBA_FIXTURE_MODULES.length))); // MODULES_COUNT
  dir.push(...vbaRec(0x0013, vbaU16(0xffff))); // PROJECTCOOKIE
  for (const m of VBA_FIXTURE_MODULES) {
    dir.push(...vbaRec(0x0019, vbaAscii(m.name))); // MODULENAME
    dir.push(...vbaRec(0x001a, vbaAscii(m.name))); // MODULESTREAMNAME
    dir.push(...vbaRec(0x0032, vbaUtf16(m.name))); // Reserved Unicode stream name
    dir.push(...vbaRec(0x0031, vbaU32(m.pcodePrefixLen))); // MODULEOFFSET
    dir.push(...vbaRec(m.document ? 0x0022 : 0x0021, [])); // MODULETYPE
    dir.push(...vbaRec(0x002b, [])); // MODULETERMINATOR
  }
  dir.push(...vbaRec(0x0010, [])); // dir Terminator — closes PROJECTMODULES, ends the dir stream
  // REFERENCEREGISTERED, appended after the terminator: real Excel files put PROJECTREFERENCES before
  // PROJECTMODULES, but the reader's uniform TLV walk doesn't care about ordering — placing it last here
  // keeps this fixture builder additive to extend rather than requiring the whole dir array reordered.
  dir.push(...vbaRec(0x000d, vbaAscii(VBA_FIXTURE_REF_MARKER)));

  const moduleStream = (m: VbaFixtureModule) => {
    const compressed = compressContainer(Uint8Array.from(m.sourceBytes));
    const out = new Uint8Array(m.pcodePrefixLen + compressed.length);
    out.set(compressed, m.pcodePrefixLen); // prefix zeros stand in for the PerformanceCache p-code
    return out;
  };

  // PROJECTwm pairs each module's MBCS name with its UTF-16 name, both NUL-terminated, ending with an
  // empty pair — one record per module, matching the dir/PROJECT streams' module list.
  const projectwm: number[] = [];
  for (const m of VBA_FIXTURE_MODULES) {
    projectwm.push(...vbaAscii(m.name), 0x00, ...vbaUtf16(m.name), 0x00, 0x00);
  }
  projectwm.push(0x00, 0x00);

  return writeCompoundFile([
    {name: 'PROJECT', data: strToU8(VBA_PROJECT_STREAM)},
    {name: 'PROJECTwm', data: Uint8Array.from(projectwm)},
    {
      name: 'VBA',
      children: [
        {name: 'dir', data: compressContainer(Uint8Array.from(dir))},
        {name: '_VBA_PROJECT', data: Uint8Array.from([0x61, 0xcc, 0x5e, 0x00, 0x00, 0x01, 0x02])},
        ...VBA_FIXTURE_MODULES.map((m) => ({name: m.name, data: moduleStream(m)})),
      ],
    },
  ]);
}

export function buildVbaFixturePackage(bin: Uint8Array): Uint8Array {
  const ms = 'http://schemas.microsoft.com/office/2006/relationships';
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
        `<Relationship Id="rId2" Type="${ms}/vbaProject" Target="vbaProject.bin"/>` +
        '</Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    ),
    'xl/vbaProject.bin': bin,
  });
}

export function vbaIndexOfBytes(haystack: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = 0; i + needle.length <= haystack.length; i++) {
    for (let j = 0; j < needle.length; j++) if (haystack[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}

// Whole-stream byte equality — used to assert a structural VBA edit leaves `_VBA_PROJECT` completely
// untouched (removeVbaModule/addVbaReference no longer reset it to the recompile cookie; Excel runs the
// project's existing p-code as-is, and resetting the cookie on a project that carries real p-code
// actively crashes the load — ADR 0019).
export function vbaBytesIdentical(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && vbaIndexOfBytes(a, b) === 0;
}

// Build a reader input of a given format family, to probe the reader's typed-error classification: a
// genuine `.xlsx` (the control that must still read), a legacy `.xls` (an OLE2/CFB compound file, via the
// production CFB writer), a binary `.xlsb` (a real ZIP whose office document is `xl/workbook.bin`),
// non-ZIP text (a CSV handed to the wrong reader), and a ZIP-headed-but-corrupt archive.
