// The default-font chain: what an unstyled cell renders in, and which of the four levels wins.
//
// The invariant under nearly every case here is that a face a *file* declared outranks anything the
// library would derive. A producer resolves the body face by script — Excel writes `等线` as font 0
// under a theme whose latin body face is Calibri — so deriving over a declaration silently rewrites
// the face of every empty cell, and with it the metric every character-unit column width means.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {INTERNAL} from './internal.ts';
import type {Font} from './style.ts';
import {Workbook} from './workbook.ts';

test('a workbook with nothing declared or authored renders in the theme body face', () => {
  const wb = new Workbook();
  assert.deepEqual(wb.defaultFont, {
    size: 11,
    color: {theme: 1},
    name: 'Calibri',
    family: 2,
    scheme: 'minor',
  });
  // Nothing was declared, and the library does not fabricate a declaration on the file's behalf.
  assert.equal(wb.declaredDefaultFont, undefined);
});

test('a declared default font is surfaced and passed through untouched', () => {
  const wb = new Workbook();
  const declared: Font = {
    size: 11,
    color: {theme: 1},
    name: 'Aptos Narrow',
    family: 2,
    scheme: 'minor',
  };
  wb[INTERNAL].restoreDefaultFont(declared);
  assert.deepEqual(wb.declaredDefaultFont, declared);
  assert.deepEqual(wb.defaultFont, declared);
});

test('a declared face outranks the theme body face it disagrees with', () => {
  // Exactly the shape Excel writes for a CJK locale: font 0 names the script face the theme's
  // <a:minorFont> nominates for Hans, while the same scheme's latin face is still Calibri.
  const wb = new Workbook();
  wb[INTERNAL].restoreDefaultFont({size: 11, color: {theme: 1}, name: '等线', scheme: 'minor'});
  assert.equal(wb.themeFonts.minor, 'Calibri');
  assert.equal(wb.defaultFont.name, '等线');
  // And its scheme claim rides through: it is the producer's own resolution, not a contradiction.
  assert.equal(wb.defaultFont.scheme, 'minor');
});

test('the theme body face reaches the default font when nothing is declared', () => {
  const wb = new Workbook();
  wb.setTheme({fonts: {minor: 'Aptos'}});
  assert.deepEqual(wb.defaultFont, {
    size: 11,
    color: {theme: 1},
    name: 'Aptos',
    family: 2,
    scheme: 'minor',
  });
});

test('an authored theme body face outranks the one a file declared', () => {
  // The read-modify-write path: without this hop, restyling a workbook's body face would reach every
  // `scheme="minor"` cell but not the unstyled ones, which is the whole complaint.
  const wb = new Workbook();
  wb[INTERNAL].restoreDefaultFont({size: 9, color: {theme: 1}, name: 'Calibri', scheme: 'minor'});
  wb.setTheme({fonts: {minor: 'Aptos'}});
  assert.equal(wb.defaultFont.name, 'Aptos');
  // The size is the workbook's, not the face's — changing the typeface says nothing about it.
  assert.equal(wb.defaultFont.size, 9);
});

test('setDefaultFont merges rather than replaces, and accumulates across calls', () => {
  const wb = new Workbook();
  wb[INTERNAL].restoreDefaultFont({size: 8, name: 'Arial'});
  wb.setDefaultFont({size: 14});
  assert.equal(wb.defaultFont.name, 'Arial', 'the declared face survives a size-only authoring');
  assert.equal(wb.defaultFont.size, 14);
  wb.setDefaultFont({bold: true});
  assert.equal(
    wb.defaultFont.size,
    14,
    'the second call adds to the first rather than resetting it',
  );
  assert.equal(wb.defaultFont.bold, true);
});

test('an authored default font outranks an authored theme body face', () => {
  const wb = new Workbook();
  wb.setTheme({fonts: {minor: 'Aptos'}});
  wb.setDefaultFont({name: 'Georgia'});
  assert.equal(wb.defaultFont.name, 'Georgia');
});

test('the resolved default font always states a size and a colour', () => {
  // A font 0 stating neither is the "missing default font" foreign readers (Numbers, and Excel in
  // some cases) warn about on open, because empty cells fall back to a default nothing defines.
  const wb = new Workbook();
  wb[INTERNAL].restoreDefaultFont({name: 'Helvetica Neue'});
  wb.setDefaultFont({italic: true});
  assert.equal(wb.defaultFont.size, 11);
  assert.deepEqual(wb.defaultFont.color, {theme: 1});
});

test('family and scheme are dropped when the resolved face is not the theme body face', () => {
  // `scheme="minor"` is a claim — that this font *is* the theme's body face — and Excel writes no
  // <scheme> at all on a font 0 naming anything else. Emitting the claim beside a contradicting
  // <name> is what makes a themed workbook render its unstyled cells in the wrong face.
  const wb = new Workbook();
  wb[INTERNAL].restoreDefaultFont({size: 11, name: 'Aptos Narrow', family: 2, scheme: 'minor'});
  wb.setDefaultFont({name: 'Georgia'});
  assert.equal(wb.defaultFont.scheme, undefined);
  assert.equal(wb.defaultFont.family, undefined);
});

test('an authored face that is the theme body face keeps the scheme claim truthful', () => {
  const wb = new Workbook();
  wb.setTheme({fonts: {minor: 'Aptos'}});
  wb.setDefaultFont({name: 'Aptos', size: 12});
  assert.equal(wb.defaultFont.scheme, 'minor');
  assert.equal(wb.defaultFont.family, 2);
});

test('a caller may state family and scheme outright', () => {
  const wb = new Workbook();
  wb.setDefaultFont({name: 'Times New Roman', family: 1});
  assert.equal(wb.defaultFont.family, 1, "the caller's word stands over the derived drop");
  assert.equal(wb.defaultFont.scheme, undefined);
});

test('an unusable default font is refused at the call that supplied it', () => {
  // Excel does not report either — it renders from some other font and never says why.
  const wb = new Workbook();
  assert.throws(() => wb.setDefaultFont({size: 0}), /positive number/);
  assert.throws(() => wb.setDefaultFont({size: Number.NaN}), /positive number/);
  assert.throws(() => wb.setDefaultFont({name: ''}), /cannot be empty/);
  // …and a refused call leaves the workbook exactly as it was.
  assert.equal(wb.defaultFont.name, 'Calibri');
});
