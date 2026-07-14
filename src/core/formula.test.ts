import assert from 'node:assert/strict';
import {test} from 'node:test';

import {mangleFormula, mangleFunctions, mangleParams, translateFormula, unmangleFunctions} from './formula.ts';

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

test('a bare-name post-2007 function (trig / bitwise / engineering) gains the prefix', () => {
  assert.equal(mangleFunctions('SEC(A1)'), '_xlfn.SEC(A1)');
  assert.equal(mangleFunctions('BITAND(5,3)'), '_xlfn.BITAND(5,3)');
  assert.equal(mangleFunctions('IMCOSH("2+i")'), '_xlfn.IMCOSH("2+i")');
  assert.equal(mangleFunctions('AGGREGATE(9,4,A1:A9)'), '_xlfn.AGGREGATE(9,4,A1:A9)');
  assert.equal(mangleFunctions('XOR(A1,B1)'), '_xlfn.XOR(A1,B1)');
  assert.equal(mangleFunctions('ISOWEEKNUM(A1)'), '_xlfn.ISOWEEKNUM(A1)');
});

test('a pre-2007 function whose name resembles a modern one is left untouched', () => {
  // SIN/COS/TAN and GAMMALN predate the frozen grammar and must NOT be prefixed, even though
  // SEC/CSC/COT and GAMMA (their newer cousins) are.
  assert.equal(mangleFunctions('SIN(A1)'), 'SIN(A1)');
  assert.equal(mangleFunctions('COS(A1)'), 'COS(A1)');
  assert.equal(mangleFunctions('GAMMALN(A1)'), 'GAMMALN(A1)');
  assert.equal(mangleFunctions('WEEKNUM(A1)'), 'WEEKNUM(A1)');
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

test('a LET parameter is _xlpm.-prefixed at its declaration and every reference', () => {
  assert.equal(mangleParams('LET(x,1,x+1)'), 'LET(_xlpm.x,1,_xlpm.x+1)');
  assert.equal(mangleParams('LET(a,B2:B9,b,2,a+b)'), 'LET(_xlpm.a,B2:B9,_xlpm.b,2,_xlpm.a+_xlpm.b)');
});

test('a LAMBDA prefixes every parameter but leaves its body cells and legacy calls alone', () => {
  assert.equal(mangleParams('LAMBDA(a,b,a+b)'), 'LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)');
  assert.equal(mangleParams('LAMBDA(r,SUM(r)+A1)'), 'LAMBDA(_xlpm.r,SUM(_xlpm.r)+A1)');
});

test('the _xlpm. prefix is scoped: a same-named reference outside the binding is untouched', () => {
  // The leading `x` is a defined-name reference, not the LET parameter — it must survive verbatim.
  assert.equal(mangleParams('x+LET(x,1,x)'), 'x+LET(_xlpm.x,1,_xlpm.x)');
  assert.equal(mangleParams('LET(x,1,x)+x'), 'LET(_xlpm.x,1,_xlpm.x)+x');
});

test('a parameter name inside a string literal is never mistaken for a reference', () => {
  assert.equal(mangleParams('LET(x,1,"x is one"&x)'), 'LET(_xlpm.x,1,"x is one"&_xlpm.x)');
});

test('a lambda-valued parameter called as a function is prefixed', () => {
  assert.equal(
    mangleParams('LET(f,LAMBDA(v,v+1),f(5))'),
    'LET(_xlpm.f,LAMBDA(_xlpm.v,_xlpm.v+1),_xlpm.f(5))',
  );
});

test('nested LET/LAMBDA scopes each bind their own parameters', () => {
  assert.equal(
    mangleParams('LET(a,B2:B9,BYROW(a,LAMBDA(r,SUM(r))))'),
    'LET(_xlpm.a,B2:B9,BYROW(_xlpm.a,LAMBDA(_xlpm.r,SUM(_xlpm.r))))',
  );
});

test('a formula with no LET/LAMBDA passes through parameter mangling unchanged', () => {
  for (const f of ['SUM(A1:A9)', 'IF(A1>0,"y","n")', 'NORM.DIST(A1,0,1,TRUE)', 'Table1[[#Data],[Col]]']) {
    assert.equal(mangleParams(f), f);
  }
});

test('mangleFormula applies both prefixes in the correct order', () => {
  assert.equal(mangleFormula('LET(x,1,x+1)'), '_xlfn.LET(_xlpm.x,1,_xlpm.x+1)');
  assert.equal(
    mangleFormula('LET(a,B2:B9,b,BYROW(a,LAMBDA(r,SUM(r))),COUNTA(UNIQUE(FILTER(a,b=1))))'),
    '_xlfn.LET(_xlpm.a,B2:B9,_xlpm.b,_xlfn.BYROW(_xlpm.a,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r))),COUNTA(_xlfn.UNIQUE(_xlfn.FILTER(_xlpm.a,_xlpm.b=1))))',
  );
});

test('mangleFormula then unmangle round-trips a LET/LAMBDA formula', () => {
  for (const f of ['LET(x,1,x+1)', 'LAMBDA(a,b,a+b)', 'LET(f,LAMBDA(v,v+1),f(5))', 'SUM(A1:A9)']) {
    assert.equal(unmangleFunctions(mangleFormula(f)), f);
  }
});

test('translateFormula shifts a relative reference by the row and column delta', () => {
  assert.equal(translateFormula('A1*2', 0, 1), 'A2*2', 'one row down');
  assert.equal(translateFormula('A1*2', 0, 2), 'A3*2', 'two rows down');
  assert.equal(translateFormula('A1', 1, 0), 'B1', 'one column across');
  assert.equal(translateFormula('B2+C3', 2, 3), 'D5+E6', 'both axes, several references');
});

test('translateFormula leaves an absolute axis fixed and shifts only the relative one', () => {
  assert.equal(translateFormula('$A$1', 3, 4), '$A$1', 'fully absolute never moves');
  assert.equal(translateFormula('$A1', 5, 1), '$A2', 'absolute column, relative row');
  assert.equal(translateFormula('A$1', 1, 5), 'B$1', 'relative column, absolute row');
  assert.equal(translateFormula('$A$1+B1', 1, 1), '$A$1+C2', 'mixed within one formula');
});

test('translateFormula is the identity for a zero delta', () => {
  assert.equal(translateFormula('SUM($A$1:B7)*C8', 0, 0), 'SUM($A$1:B7)*C8');
});

test('translateFormula shifts both endpoints of a range independently', () => {
  assert.equal(translateFormula('SUM(A1:B2)', 1, 10), 'SUM(B11:C12)');
  assert.equal(translateFormula('A1:$B$2', 0, 5), 'A6:$B$2', 'the absolute endpoint stays');
});

test('translateFormula never touches a function name or a defined name', () => {
  assert.equal(translateFormula('SUM(A1:A3)', 0, 1), 'SUM(A2:A4)', 'SUM has no row digits');
  assert.equal(translateFormula('TaxRate*A1', 2, 2), 'TaxRate*C3', 'a defined name is left alone');
  assert.equal(translateFormula('LOG10(A1)', 0, 1), 'LOG10(A2)', 'a call ending in digits is not a reference');
});

test('translateFormula shifts a sheet-qualified cell but not the sheet name', () => {
  assert.equal(translateFormula('Sheet1!A1', 0, 1), 'Sheet1!A2', 'the cell after ! moves');
  assert.equal(translateFormula('Q1!A1', 0, 1), 'Q1!A2', 'a sheet name that looks like a reference is untouched');
  assert.equal(translateFormula("'My Sheet'!A1+B2", 1, 1), "'My Sheet'!B2+C3", 'a quoted sheet name is copied verbatim');
});

test('translateFormula copies a string literal verbatim, references outside it still move', () => {
  assert.equal(translateFormula('IF(A1>0,"A1 is B2",B2)', 0, 1), 'IF(A2>0,"A1 is B2",B3)');
});
