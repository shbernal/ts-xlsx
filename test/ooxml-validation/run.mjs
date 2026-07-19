#!/usr/bin/env node

import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {mkdtemp, readFile, rm, writeFile} from 'node:fs/promises';
import {createRequire} from 'node:module';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const require = createRequire(import.meta.url);
const ExcelJS = require('../../lib/exceljs.nodejs.js');
const JSZip = require('jszip');

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..', '..');
const PROJECT = path.join(ROOT, 'tools', 'ooxml-validator', 'OoxmlValidator.csproj');
const BASELINE = JSON.parse(await readFile(path.join(HERE, 'allowed-errors.json'), 'utf8'));
const TIMEOUT_MS = 30_000;
const MAX_OUTPUT_BYTES = 16 * 1024 * 1024;

function runDotnet(files, {format = 'Microsoft365'} = {}) {
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

  return new Promise((resolve, reject) => {
    const child = spawn('dotnet', args, {cwd: ROOT, stdio: ['ignore', 'pipe', 'pipe']});
    const stdout = [];
    const stderr = [];
    let outputBytes = 0;
    let settled = false;
    const timeout = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`OOXML validator timed out after ${TIMEOUT_MS}ms`));
    }, TIMEOUT_MS);

    const collect = (target) => (chunk) => {
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

async function writeBufferedWorkbook(file) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.columns = [
    {header: 'Name', key: 'name', width: 20},
    {header: 'Value', key: 'value', width: 12},
  ];
  sheet.addRow({name: 'alpha', value: 42});
  sheet.getCell('A2').font = {bold: true, color: {argb: 'FF336699'}};
  sheet.getCell('B2').dataValidation = {type: 'whole', operator: 'between', formulae: [0, 100]};
  sheet.getCell('B3').value = {formula: 'SUM(B2:B2)', result: 42};
  sheet.addTable({
    name: 'DataTable',
    ref: 'D1',
    headerRow: true,
    columns: [{name: 'Label'}, {name: 'Amount'}],
    rows: [['alpha', 42]],
  });
  await workbook.xlsx.writeFile(file);
}

async function writeStreamingWorkbook(file) {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: file,
    useSharedStrings: true,
    useStyles: true,
  });
  const sheet = workbook.addWorksheet('Stream');
  sheet.columns = [
    {header: 'Name', key: 'name', width: 20},
    {header: 'Value', key: 'value', width: 12},
  ];
  sheet.addRow({name: 'alpha', value: 42}).commit();
  sheet.addRow({name: 'beta', value: 7}).commit();
  sheet.commit();
  await workbook.commit();
}

async function rewritePackage(source, destination, transform) {
  const zip = await JSZip.loadAsync(await readFile(source));
  await transform(zip);
  await writeFile(destination, await zip.generateAsync({type: 'nodebuffer'}));
}

async function makeSchemaCleanControl(source, destination) {
  await rewritePackage(source, destination, async (zip) => {
    const stylesPart = zip.file('xl/styles.xml');
    assert.ok(stylesPart, 'generated workbook must contain xl/styles.xml');
    const styles = await stylesPart.async('string');
    const corrected = styles.replace(
      '<font><color theme="1"/><family val="2"/><scheme val="minor"/><sz val="11"/><name val="Calibri"/></font>',
      '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
    );
    assert.notStrictEqual(corrected, styles, 'known font-order sequence must be present');
    zip.file('xl/styles.xml', corrected);
  });
}

async function makeSchemaInvalidControl(source, destination) {
  await rewritePackage(source, destination, async (zip) => {
    const sheetPart = zip.file('xl/worksheets/sheet1.xml');
    assert.ok(sheetPart, 'generated workbook must contain xl/worksheets/sheet1.xml');
    const xml = await sheetPart.async('string');
    const invalid = xml.replace('</worksheet>', '<unexpectedValidatorProbe/></worksheet>');
    assert.notStrictEqual(invalid, xml, 'worksheet close tag must be present');
    zip.file('xl/worksheets/sheet1.xml', invalid);
  });
}

function fingerprint(error) {
  return {
    id: error.id,
    type: error.type,
    partUri: error.partUri,
    xpath: error.xpath,
  };
}

function parseReport(result, expectedCode) {
  assert.strictEqual(
    result.code,
    expectedCode,
    `validator exit code; stderr=${result.stderr}; stdout=${result.stdout}`,
  );
  assert.doesNotThrow(
    () => JSON.parse(result.stdout),
    `validator must emit JSON: ${result.stdout}`,
  );
  return JSON.parse(result.stdout);
}

async function main() {
  const temp = await mkdtemp(path.join(tmpdir(), 'ts-xlsx-ooxml-'));
  try {
    const buffered = path.join(temp, 'buffered.xlsx');
    const streaming = path.join(temp, 'streaming.xlsx');
    const clean = path.join(temp, 'clean.xlsx');
    const invalid = path.join(temp, 'invalid.xlsx');
    const truncated = path.join(temp, 'truncated.xlsx');
    const unsupported = path.join(temp, 'unsupported.txt');

    await writeBufferedWorkbook(buffered);
    await writeStreamingWorkbook(streaming);
    await makeSchemaCleanControl(buffered, clean);
    await makeSchemaInvalidControl(clean, invalid);
    await writeFile(truncated, (await readFile(clean)).subarray(0, 128));
    await writeFile(unsupported, 'not an xlsx');

    const report = parseReport(
      await runDotnet([buffered, streaming, clean, invalid, truncated]),
      1,
    );
    assert.strictEqual(report.format, 'Microsoft365');
    const byName = new Map(report.results.map((result) => [path.basename(result.file), result]));

    for (const [name, expected] of Object.entries(BASELINE)) {
      const result = byName.get(name);
      assert.ok(result, `missing validator result for ${name}`);
      assert.deepStrictEqual(result.errors.map(fingerprint), expected, `${name} baseline changed`);
      assert.strictEqual(
        result.valid,
        false,
        `${name} must remain baselined until the writer is fixed`,
      );
    }

    assert.deepStrictEqual(
      byName.get('clean.xlsx')?.errors,
      [],
      'schema-clean control must validate',
    );
    assert.strictEqual(byName.get('clean.xlsx')?.valid, true);

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

    parseReport(await runDotnet([clean]), 0);

    const badInvocation = await runDotnet([unsupported]);
    assert.strictEqual(badInvocation.code, 2);
    assert.match(badInvocation.stderr, /Only \.xlsx files are supported/);
    assert.strictEqual(badInvocation.stdout, '');

    console.log('ooxml validation: buffered + streaming baselines and clean/error controls passed');
  } finally {
    await rm(temp, {recursive: true, force: true});
  }
}

main().catch((error) => {
  console.error(error.stack ?? error);
  process.exitCode = 1;
});
