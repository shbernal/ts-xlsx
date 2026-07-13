import assert from 'node:assert/strict';
import {test} from 'node:test';

import {mangleFunctions, unmangleFunctions} from './formula.ts';

test('a modern function called by its plain name gains the _xlfn. prefix', () => {
  assert.equal(mangleFunctions('FILTER(B1:D1,B2:D2=1)'), '_xlfn.FILTER(B1:D1,B2:D2=1)');
  assert.equal(mangleFunctions('XLOOKUP(1,B:B,C:C)'), '_xlfn.XLOOKUP(1,B:B,C:C)');
});

test('a legacy function is left untouched', () => {
  assert.equal(mangleFunctions('SUM(A1:A9)'), 'SUM(A1:A9)');
  assert.equal(mangleFunctions('IF(A1>0,"y","n")'), 'IF(A1>0,"y","n")');
});

test('nested modern functions each get the prefix, legacy ones do not', () => {
  assert.equal(
    mangleFunctions('SUM(FILTER(A:A,B:B=1))'),
    'SUM(_xlfn.FILTER(A:A,B:B=1))',
  );
  assert.equal(
    mangleFunctions('COUNTA(UNIQUE(FILTER(a,b=1)))'),
    'COUNTA(_xlfn.UNIQUE(_xlfn.FILTER(a,b=1)))',
  );
});

test('an already-prefixed function is not double-prefixed', () => {
  assert.equal(mangleFunctions('_xlfn.XLOOKUP(1,B:B,C:C)'), '_xlfn.XLOOKUP(1,B:B,C:C)');
  assert.ok(!mangleFunctions('_xlfn.XLOOKUP(1,B:B,C:C)').includes('_xlfn._xlfn'));
});

test('a function name inside a string literal is never mangled', () => {
  assert.equal(mangleFunctions('IF(A1="FILTER(",1,2)'), 'IF(A1="FILTER(",1,2)');
  assert.equal(mangleFunctions('CONCAT("SORT()",A1)'), '_xlfn.CONCAT("SORT()",A1)');
});

test('mangling introduces no @ implicit-intersection operator', () => {
  const out = mangleFunctions('IFS(B1>0,"pos",B1<0,"neg")');
  assert.equal(out, '_xlfn.IFS(B1>0,"pos",B1<0,"neg")');
  assert.ok(!/(^|[^A-Za-z0-9_])@/.test(out));
});

test('matching is case-insensitive on the function name but preserves its casing', () => {
  assert.equal(mangleFunctions('filter(A:A,B:B=1)'), '_xlfn.filter(A:A,B:B=1)');
  assert.equal(mangleFunctions('Filter(A:A,B:B=1)'), '_xlfn.Filter(A:A,B:B=1)');
});

test('a LET/LAMBDA formula gets the _xlfn. prefix on every modern function', () => {
  const out = mangleFunctions('LET(a,B2:B9,b,BYROW(a,LAMBDA(r,SUM(r))),COUNTA(UNIQUE(FILTER(a,b=1))))');
  assert.ok(out.includes('_xlfn.LET'));
  assert.ok(out.includes('_xlfn.BYROW'));
  assert.ok(out.includes('_xlfn.LAMBDA'));
  assert.ok(out.includes('_xlfn.UNIQUE'));
  assert.ok(out.includes('_xlfn.FILTER'));
  assert.ok(out.includes('SUM(r)'));
  assert.ok(out.includes('COUNTA('));
});

test('a dotted 2010 statistical function is matched whole and prefixed', () => {
  assert.equal(mangleFunctions('NORM.DIST(A1,0,1,TRUE)'), '_xlfn.NORM.DIST(A1,0,1,TRUE)');
  assert.equal(mangleFunctions('BETA.INV(0.5,2,3)'), '_xlfn.BETA.INV(0.5,2,3)');
  assert.equal(mangleFunctions('T.DIST.2T(2,10)'), '_xlfn.T.DIST.2T(2,10)');
  assert.equal(mangleFunctions('NORM.S.INV(0.9)'), '_xlfn.NORM.S.INV(0.9)');
});

test('a dotted function only gets one prefix, on the whole name, not per segment', () => {
  const out = mangleFunctions('CHISQ.DIST.RT(3,2)');
  assert.equal(out, '_xlfn.CHISQ.DIST.RT(3,2)');
  assert.ok(!out.includes('_xlfn.DIST'), 'the tail segment must not be prefixed on its own');
});

test('an already-prefixed dotted function is not double-prefixed', () => {
  assert.equal(mangleFunctions('_xlfn.NORM.DIST(A1,0,1,TRUE)'), '_xlfn.NORM.DIST(A1,0,1,TRUE)');
  assert.ok(!mangleFunctions('_xlfn.NORM.DIST(A1,0,1,TRUE)').includes('_xlfn._xlfn'));
});

test('a dotted function nested beside plain and legacy functions is prefixed correctly', () => {
  assert.equal(
    mangleFunctions('SUM(NORM.DIST(A1,0,1,TRUE),STDEV.S(B:B),AVERAGE(C:C))'),
    'SUM(_xlfn.NORM.DIST(A1,0,1,TRUE),_xlfn.STDEV.S(B:B),AVERAGE(C:C))',
  );
});

test('a decimal literal adjacent to a dotted call is not mistaken for a function', () => {
  assert.equal(mangleFunctions('NORM.DIST(A1,0,1,TRUE)*1.5'), '_xlfn.NORM.DIST(A1,0,1,TRUE)*1.5');
});

test('unmangle strips _xlfn. and _xlpm. back to the plain names', () => {
  assert.equal(unmangleFunctions('_xlfn.XLOOKUP(1,B:B,C:C)'), 'XLOOKUP(1,B:B,C:C)');
  assert.equal(unmangleFunctions('_xlfn.LET(_xlpm.a,B2:B9,_xlpm.a)'), 'LET(a,B2:B9,a)');
  assert.equal(unmangleFunctions('_xlfn.NORM.DIST(A1,0,1,TRUE)'), 'NORM.DIST(A1,0,1,TRUE)');
});

test('mangle then unmangle round-trips a plain formula', () => {
  for (const f of ['FILTER(A:A,B:B=1)', 'SUM(A1:A9)', 'IFS(B1>0,"pos")', 'XLOOKUP(1,B:B,C:C)', 'NORM.DIST(A1,0,1,TRUE)', 'T.DIST.2T(2,10)']) {
    assert.equal(unmangleFunctions(mangleFunctions(f)), f);
  }
});
