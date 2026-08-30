// The package plumbing every xlsx test needs: unzip a written workbook, reach one part, or feed the
// reader hand-authored markup.
//
// This existed nineteen times across fifteen files, in four families, with three different answers
// to "the part is not there": a `TypeError` from `strFromU8` on undefined, a silent empty string, or
// an assertion naming the part. The silent one was the common spelling in the two files with the
// most negative assertions, where a dozen `assert.doesNotMatch` calls would have gone on passing
// against `''` had the writer ever renamed the part they read. So every accessor here asserts,
// except the one whose whole purpose is absence.
//
// Not a `.test.ts` file: `node --test` would try to run it and find no tests. `tsconfig.build.json`
// excludes the `.test-support.ts` suffix, so it ships nowhere.

import assert from 'node:assert/strict';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

/** The first worksheet's part path, which most writer tests are reaching for. */
export const SHEET1 = 'xl/worksheets/sheet1.xml';

/** Every part of a package, decoded to text, keyed by part path. */
export function partsOf(pkg: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(unzipSync(pkg))) out[name] = strFromU8(bytes);
  return out;
}

/** One named part's text. Fails the test, naming the part, when it is absent. */
export function partText(pkg: Uint8Array, name: string): string {
  const bytes = unzipSync(pkg)[name];
  assert.ok(bytes, `expected part ${name}`);
  return strFromU8(bytes);
}

/** One named part's text, or undefined: for a test asserting a part was *not* written. */
export function optionalPartText(pkg: Uint8Array, name: string): string | undefined {
  const bytes = unzipSync(pkg)[name];
  return bytes === undefined ? undefined : strFromU8(bytes);
}

/** The first worksheet's XML. */
export function sheetXml(pkg: Uint8Array): string {
  return partText(pkg, SHEET1);
}

/** Write a workbook and read it straight back: the round-trip under test. */
export function roundtrip(workbook: Workbook): Workbook {
  return readXlsx(writeXlsx(workbook));
}

/**
 * Read an xlsx built from hand-authored parts patched into a written package, keyed by part path, so
 * a case can feed the reader markup the writer itself only produces on round-trip: an Excel-authored
 * x14 extLst block, a foreign dxf table, a docProps element left empty.
 *
 * The base package is a one-sheet workbook with `A1` set, so a test can assert the read succeeded at
 * all rather than passing by reading nothing.
 */
export function readPatched(parts: Record<string, string>): Workbook {
  const base = new Workbook();
  base.addWorksheet('S').getCell('A1').value = 1;
  const files = unzipSync(writeXlsx(base));
  for (const [name, xml] of Object.entries(parts)) files[name] = strToU8(xml);
  return readXlsx(zipSync(files));
}
