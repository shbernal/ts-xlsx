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

// One unzip per package, not one per accessor. Every accessor below used to inflate the whole
// package to answer for a single part, and a test that checks four parts of one workbook paid for
// four; `read.test.ts` does exactly that repeatedly. Keyed on the package's own bytes, so a test
// that writes twice gets two entries and a test that never re-reads pays nothing, and weakly so
// that a suite building hundreds of workbooks does not hold them all.
//
// The record `unzipSync` returns is the memo itself, so nothing may write to it. `patchParts` is
// the one caller that wants to, and copies first.
const inflated = new WeakMap<Uint8Array, Record<string, Uint8Array>>();

function filesOf(pkg: Uint8Array): Record<string, Uint8Array> {
  const hit = inflated.get(pkg);
  if (hit !== undefined) return hit;
  const files = unzipSync(pkg);
  inflated.set(pkg, files);
  return files;
}

/** Every part of a package, decoded to text, keyed by part path. */
export function partsOf(pkg: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(filesOf(pkg))) out[name] = strFromU8(bytes);
  return out;
}

/**
 * Every part of the package a workbook writes to, decoded to text: {@link partsOf} for the many tests
 * that start from a `Workbook` rather than from bytes.
 *
 * Two test files declared this locally, byte-identical down to the comment, each shadowing the
 * imported `partsOf` it wrapped, which is the drift this module exists to stop.
 */
export function partsWritten(workbook: Workbook): Record<string, string> {
  return partsOf(writeXlsx(workbook));
}

/** One named part's text. Fails the test, naming the part, when it is absent. */
export function partText(pkg: Uint8Array, name: string): string {
  const bytes = filesOf(pkg)[name];
  assert.ok(bytes, `expected part ${name}`);
  return strFromU8(bytes);
}

/**
 * One named part's text out of an already-decoded {@link partsOf}/{@link partsWritten} record. Fails
 * the test naming the part when it is absent, which {@link partText} does for a package's bytes: a
 * test that already has the whole record should not have to unzip again to get the assertion.
 */
export function partIn(parts: Record<string, string>, name: string): string {
  const xml = parts[name];
  assert.ok(xml !== undefined, `expected part ${name}`);
  return xml;
}

/** One named part's text, or undefined: for a test asserting a part was *not* written. */
export function optionalPartText(pkg: Uint8Array, name: string): string | undefined {
  const bytes = filesOf(pkg)[name];
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
 * Rewrite named parts of a package and return the new bytes, asserting each named part exists first.
 *
 * The other half of what {@link readPatched} does, for a test that patches a package it built itself
 * rather than the stock one, or that wants the bytes rather than the model. Every site doing this by
 * hand read the part it was about to patch through a cast or a `?? new Uint8Array()`, which is the
 * failure this module exists to remove: the first spelling throws a `TypeError` naming nothing when
 * the writer renames a part, and the second silently patches an empty string, after which every
 * negative assertion built on the result passes for the wrong reason.
 */
export function patchParts(
  pkg: Uint8Array,
  edits: Record<string, (xml: string) => string>,
): Uint8Array {
  const files = {...filesOf(pkg)};
  for (const [name, edit] of Object.entries(edits)) {
    const bytes = files[name];
    assert.ok(bytes, `expected part ${name}`);
    files[name] = strToU8(edit(strFromU8(bytes)));
  }
  return zipSync(files);
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
