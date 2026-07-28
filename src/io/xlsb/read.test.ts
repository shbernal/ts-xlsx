import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import path from 'node:path';
import {test} from 'node:test';
import {fileURLToPath} from 'node:url';

import {strToU8, unzipSync, zipSync} from 'fflate';

import {UnsupportedFormatError} from '../opc/errors.ts';
import {readXlsx} from '../xlsx/read.ts';
import {XlsbParseError} from './errors.ts';
import {readXlsb} from './read.ts';

// The corpus owns the implementation-blind "reads like its XML twin" property; these are the
// white-box checks that go with the modules — the wiring, and the failure modes a fixture cannot show.
const FIXTURE = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '../../../test/corpus/fixtures/xlsb-binary-workbook-reads-like-its-xlsx-twin/source.xlsb',
);

const fixture = (): Uint8Array => readFileSync(FIXTURE);

test('readXlsb reads a binary package into the workbook model', () => {
  const workbook = readXlsb(fixture());
  assert.deepEqual(
    workbook.worksheets.map((sheet) => sheet.name),
    ['Quiet', 'Grid', 'Values'],
  );
  assert.equal(workbook.getWorksheet('Values')?.getCell('B2').value, 10);
});

test('readXlsx auto-detects a binary package and produces the same workbook readXlsb does', () => {
  // The auto-detect path must not be a second, subtly different reader — it is the same codec, handed
  // an already-inflated package.
  const auto = readXlsx(fixture());
  const explicit = readXlsb(fixture());
  const project = (workbook: ReturnType<typeof readXlsb>): string =>
    JSON.stringify(workbook.worksheets.map((sheet) => sheet.model));
  assert.equal(project(auto), project(explicit));
});

test('readXlsb rejects an XML .xlsx package rather than reading it as empty', () => {
  const archive = zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
    'xl/workbook.xml': strToU8('<workbook/>'),
  });
  assert.throws(
    () => readXlsb(archive),
    (error: unknown) => {
      assert.ok(error instanceof UnsupportedFormatError);
      assert.match(error.message, /workbook\.bin/);
      return true;
    },
  );
});

test('a binary workbook part with a lying record length fails closed', () => {
  const archive = zipSync({
    'xl/workbook.bin': Uint8Array.of(0x81, 0x00, 0x02, 0xff, 0xff, 0xff, 0x7f, 0x01),
  });
  assert.throws(() => readXlsb(archive), XlsbParseError);
});

test('a sheet whose relationship target is missing yields an empty sheet, not a crash', () => {
  // A dangling reference is a damaged file, not a hostile one: the workbook still names the sheet, so
  // the model keeps it rather than dropping the sheet the caller can see in Excel's tab bar.
  const workbook = readXlsb(withoutPart('xl/worksheets/sheet3.bin'));
  assert.equal(workbook.worksheets.length, 3);
  assert.equal(workbook.getWorksheet('Values')?.model.cells.length, 0);
});

test('a package with no style sheet reads its values with every cell unstyled', () => {
  const workbook = readXlsb(withoutPart('xl/styles.bin'));
  const cell = workbook.getWorksheet('Values')?.getCell('A1');
  assert.equal(cell?.value, 'kind');
  assert.equal(cell?.font, undefined);
  assert.equal(cell?.fill, undefined);
});

test('a package with no shared-string table reads its pooled cells as empty strings', () => {
  const workbook = readXlsb(withoutPart('xl/sharedStrings.bin'));
  // The value is gone with the pool, but the cell — and the rest of the sheet — still reads.
  assert.equal(workbook.getWorksheet('Values')?.getCell('A1').value, '');
  assert.equal(workbook.getWorksheet('Values')?.getCell('B2').value, 10);
});

// Rebuild the fixture archive without one part, to exercise the reader's tolerance of a damaged
// package without committing a second fixture for each variant.
function withoutPart(dropped: string): Uint8Array {
  const parts = unzipSync(fixture());
  delete parts[dropped];
  return zipSync(parts);
}
