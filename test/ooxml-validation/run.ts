#!/usr/bin/env node

import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {mkdtemp, readFile, rm, writeFile} from 'node:fs/promises';
import {createRequire} from 'node:module';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import type JSZipType from 'jszip';
import {Workbook} from '../../src/core/workbook.ts';
import {writeXlsx} from '../../src/io/xlsx/write.ts';
import {WorkbookStreamWriter} from '../../src/io/xlsx/write-stream.ts';

/** Structured diagnostic emitted per validation problem by OpenXmlValidator. */
interface ValidationError {
  readonly id: string;
  readonly type: string;
  readonly partUri: string;
  readonly xpath: string;
}

/** The stable subset of a diagnostic used to detect baseline drift. */
type ValidationFingerprint = Pick<ValidationError, 'id' | 'type' | 'partUri' | 'xpath'>;

/** Per-file validation outcome inside a report. */
interface ValidationResult {
  readonly file: string;
  readonly valid: boolean;
  readonly errors: readonly ValidationError[];
}

/** The full JSON document the validator prints to stdout. */
interface ValidationReport {
  readonly format: string;
  readonly results: readonly ValidationResult[];
}

/** Baselined-until-fixed diagnostics, keyed by workbook basename. Empty while the writer is clean —
 * an entry is a *known-open* writer bug we've chosen to track, never a mute button for a new one. */
type Baseline = Readonly<Record<string, readonly ValidationFingerprint[]>>;

/** The captured outcome of one `dotnet run` invocation. */
interface DotnetRun {
  readonly code: number | null;
  readonly stdout: string;
  readonly stderr: string;
}

const require = createRequire(import.meta.url);
const JSZip = require('jszip') as typeof JSZipType;

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..', '..');
const PROJECT = path.join(ROOT, 'tools', 'ooxml-validator', 'OoxmlValidator.csproj');
const BASELINE = JSON.parse(
  await readFile(path.join(HERE, 'allowed-errors.json'), 'utf8'),
) as Baseline;
const TIMEOUT_MS = 30_000;
const MAX_OUTPUT_BYTES = 16 * 1024 * 1024;

// The buffered and both streaming outputs are the packages under test: every one must validate against
// the frozen baseline (empty today, so: clean). A new diagnostic on any of them fails the gate.
const WRITER_FILES = ['buffered.xlsx', 'streaming-inline.xlsx', 'streaming-shared.xlsx'] as const;

function runDotnet(
  files: readonly string[],
  {format = 'Microsoft365'}: {format?: string} = {},
): Promise<DotnetRun> {
  const args = [
    'run',
    '--project',
    PROJECT,
    '--configuration',
    'Release',
    '--no-restore',
    '--',
    '--format',
    format,
    ...files,
  ];

  return new Promise<DotnetRun>((resolve, reject) => {
    const child = spawn('dotnet', args, {
      cwd: ROOT,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    const stdout: Buffer[] = [];
    const stderr: Buffer[] = [];
    let outputBytes = 0;
    let settled = false;
    const timeout = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`OOXML validator timed out after ${TIMEOUT_MS}ms`));
    }, TIMEOUT_MS);

    const collect = (target: Buffer[]) => (chunk: Buffer) => {
      outputBytes += chunk.length;
      if (outputBytes > MAX_OUTPUT_BYTES) {
        child.kill('SIGKILL');
        reject(new Error(`OOXML validator output exceeded ${MAX_OUTPUT_BYTES} bytes`));
        return;
      }
      target.push(chunk);
    };
    child.stdout.on('data', collect(stdout));
    child.stderr.on('data', collect(stderr));
    child.on('error', (error) => {
      clearTimeout(timeout);
      if (!settled) {
        settled = true;
        reject(error);
      }
    });
    child.on('close', (code) => {
      clearTimeout(timeout);
      if (settled) return;
      settled = true;
      resolve({
        code,
        stdout: Buffer.concat(stdout).toString('utf8'),
        stderr: Buffer.concat(stderr).toString('utf8'),
      });
    });
  });
}

// Exercise a representative slice of the buffered writer — styled font, data validation, a formula, and
// a table over its own cells with a totals row carrying a custom <totalsRowFormula> — so the oracle sees
// more than a bare grid and validates the totals-row markup against the schema.
async function writeBufferedWorkbook(file: string): Promise<void> {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Data');
  Object.assign(sheet.getColumn(1), {key: 'name', width: 20});
  Object.assign(sheet.getColumn(2), {key: 'value', width: 12});
  sheet.addRow(['Name', 'Value']);
  sheet.addRow(['alpha', 42]);
  sheet.getCell('A2').font = {bold: true, color: {argb: 'FF336699'}};
  sheet.addDataValidation('B2:B20', {type: 'whole', operator: 'between', formulae: [0, 100]});
  sheet.getCell('B3').value = {formula: 'SUM(B2:B2)', result: 42};
  sheet.getCell('D1').value = 'Label';
  sheet.getCell('E1').value = 'Amount';
  sheet.getCell('D2').value = 'alpha';
  sheet.getCell('E2').value = 42;
  // The totals row's cells (label + custom SUBTOTAL-less formula) are materialised by addTable, so the
  // oracle checks both the <totalsRowFormula> child and the grid cells it writes.
  sheet.addTable({
    name: 'DataTable',
    ref: 'D1',
    headerRow: true,
    totalsRow: true,
    columns: [
      {name: 'Label', totalsRowLabel: 'Total'},
      {name: 'Amount', totalsRowFunction: 'custom', totalsRowFormula: 'SUM(DataTable[Amount])*1.1'},
    ],
    rowCount: 1,
  });
  await writeFile(file, writeXlsx(workbook));
}

// The streaming writer must be clean in both string-storage modes: inline (eager per-row flush) and
// shared-strings (whole-workbook pool). They travel different serialisation paths, so both are exercised.
async function writeStreamingWorkbook(file: string, useSharedStrings: boolean): Promise<void> {
  const writer = new WorkbookStreamWriter({useSharedStrings});
  const sheet = writer.addWorksheet('Stream');
  sheet.addRow(['Name', 'Value']).commit();
  sheet.addRow(['alpha', 42]).commit();
  sheet.addRow(['beta', 7]).commit();
  sheet.commit();
  await writeFile(file, await writer.commit());
}

async function rewritePackage(
  source: string,
  destination: string,
  transform: (zip: JSZipType) => Promise<void>,
): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(source));
  await transform(zip);
  await writeFile(destination, await zip.generateAsync({type: 'nodebuffer'}));
}

// A negative control: inject an element the worksheet schema forbids, so a passing run proves the oracle
// still discriminates rather than rubber-stamping. Derived from a known-clean package.
async function makeSchemaInvalidControl(source: string, destination: string): Promise<void> {
  await rewritePackage(source, destination, async (zip) => {
    const sheetPart = zip.file('xl/worksheets/sheet1.xml');
    assert.ok(sheetPart, 'generated workbook must contain xl/worksheets/sheet1.xml');
    const xml = await sheetPart.async('string');
    const invalid = xml.replace('</worksheet>', '<unexpectedValidatorProbe/></worksheet>');
    assert.notStrictEqual(invalid, xml, 'worksheet close tag must be present');
    zip.file('xl/worksheets/sheet1.xml', invalid);
  });
}

function fingerprint(error: ValidationError): ValidationFingerprint {
  return {
    id: error.id,
    type: error.type,
    partUri: error.partUri,
    xpath: error.xpath,
  };
}

function parseReport(result: DotnetRun, expectedCode: number): ValidationReport {
  assert.strictEqual(
    result.code,
    expectedCode,
    `validator exit code; stderr=${result.stderr}; stdout=${result.stdout}`,
  );
  assert.doesNotThrow(
    () => JSON.parse(result.stdout),
    `validator must emit JSON: ${result.stdout}`,
  );
  return JSON.parse(result.stdout) as ValidationReport;
}

async function main(): Promise<void> {
  const temp = await mkdtemp(path.join(tmpdir(), 'ts-xlsx-ooxml-'));
  try {
    const at = (name: string) => path.join(temp, name);
    const invalid = at('invalid.xlsx');
    const truncated = at('truncated.xlsx');
    const unsupported = at('unsupported.txt');

    await writeBufferedWorkbook(at('buffered.xlsx'));
    await writeStreamingWorkbook(at('streaming-inline.xlsx'), false);
    await writeStreamingWorkbook(at('streaming-shared.xlsx'), true);
    await makeSchemaInvalidControl(at('buffered.xlsx'), invalid);
    await writeFile(truncated, (await readFile(at('buffered.xlsx'))).subarray(0, 128));
    await writeFile(unsupported, 'not an xlsx');

    // Both negative controls make this a non-zero run; the writer files are asserted clean below.
    const report = parseReport(await runDotnet([...WRITER_FILES.map(at), invalid, truncated]), 1);
    assert.strictEqual(report.format, 'Microsoft365');
    const byName = new Map(report.results.map((result) => [path.basename(result.file), result]));

    for (const name of WRITER_FILES) {
      const result = byName.get(name);
      assert.ok(result, `missing validator result for ${name}`);
      const expected = BASELINE[name] ?? [];
      assert.deepStrictEqual(
        result.errors.map(fingerprint),
        expected,
        `${name} diverged from its baseline — fix the writer, do not baseline a new error`,
      );
      assert.strictEqual(
        result.valid,
        expected.length === 0,
        `${name} validity must match its (empty) baseline`,
      );
    }

    const invalidResult = byName.get('invalid.xlsx');
    assert.strictEqual(invalidResult?.valid, false);
    assert.ok(
      invalidResult?.errors.some(
        (error) => error.type === 'Schema' && error.partUri === '/xl/worksheets/sheet1.xml',
      ),
      'invalid worksheet must produce a structured schema diagnostic',
    );

    const truncatedResult = byName.get('truncated.xlsx');
    assert.strictEqual(truncatedResult?.valid, false);
    assert.deepStrictEqual(
      truncatedResult?.errors.map((error) => error.id),
      ['PackageOpenError'],
    );
    assert.deepStrictEqual(
      truncatedResult?.errors.map((error) => error.type),
      ['Package'],
    );

    // The clean exit path: the writer outputs alone must return exit 0.
    parseReport(await runDotnet(WRITER_FILES.map(at)), 0);

    const badInvocation = await runDotnet([unsupported]);
    assert.strictEqual(badInvocation.code, 2);
    assert.match(badInvocation.stderr, /Only \.xlsx files are supported/);
    assert.strictEqual(badInvocation.stdout, '');

    console.log('ooxml validation: buffered + streaming outputs clean; error controls detected');
  } finally {
    await rm(temp, {recursive: true, force: true});
  }
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? (error.stack ?? error.message) : error);
  process.exitCode = 1;
});
