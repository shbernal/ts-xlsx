import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XlsbParseError} from './errors.ts';
import {decodeFormula, type FormulaScope, formulaAnchor} from './formula.ts';
import {fixedArityFor, functionNameFor} from './ptg-functions.ts';

function bytes(...values: number[]): Uint8Array {
  return Uint8Array.from(values);
}

const NONE = new Uint8Array(0);

// A workbook of three sheets whose extern-sheet table names each of them and, at index 3, an
// external workbook — the shape the resolver has to tell apart.
const SCOPE: FormulaScope = {
  sheetNames: ['Calc', 'Data', 'More'],
  externSheets: [
    {supBook: 0, firstSheet: 0, lastSheet: 0},
    {supBook: 0, firstSheet: 1, lastSheet: 2},
    {supBook: 0, firstSheet: 9, lastSheet: 9},
    {supBook: 1, firstSheet: 0, lastSheet: 0},
  ],
  selfSupBook: 0,
  names: ['Rate'],
};

// `PtgRef` (value class): the ptg, a 4-byte row, then the packed column word.
function ref(row: number, column: number, flags = 0xc000): Uint8Array {
  const out = new Uint8Array(7);
  const view = new DataView(out.buffer);
  view.setUint8(0, 0x44);
  view.setUint32(1, row, true);
  view.setUint16(5, column | flags, true);
  return out;
}

test('a postfix stream folds back into infix text', () => {
  // 1, 2, 3, multiply, add — the tokens behind `1+2*3`.
  const rgce = bytes(0x1e, 1, 0, 0x1e, 2, 0, 0x1e, 3, 0, 0x05, 0x03);
  assert.equal(decodeFormula(rgce, NONE, SCOPE), '1+2*3');
});

test('a stored parenthesis is emitted where the author typed it', () => {
  const rgce = bytes(0x1e, 1, 0, 0x1e, 2, 0, 0x03, 0x15, 0x1e, 3, 0, 0x05);
  assert.equal(decodeFormula(rgce, NONE, SCOPE), '(1+2)*3');
});

test('the relative-axis flags decide where the dollar signs go', () => {
  assert.equal(decodeFormula(ref(0, 0, 0x0000), NONE, SCOPE), '$A$1');
  assert.equal(decodeFormula(ref(1, 2, 0x4000), NONE, SCOPE), 'C$2');
  assert.equal(decodeFormula(ref(1, 2, 0x8000), NONE, SCOPE), '$C2');
  assert.equal(decodeFormula(ref(1, 2, 0xc000), NONE, SCOPE), 'C2');
});

test('a 3-D reference resolves its sheet span through the extern-sheet table', () => {
  // PtgRef3d (value class): ptg, ixti, row, packed column.
  const at = (ixti: number) => bytes(0x5a, ixti, 0, 0, 0, 0, 0, 0x00, 0xc0);
  assert.equal(decodeFormula(at(0), NONE, SCOPE), 'Calc!A1');
  assert.equal(decodeFormula(at(1), NONE, SCOPE), 'Data:More!A1');
  // A span naming a sheet the workbook does not have resolves to nothing rather than to a guess.
  assert.equal(decodeFormula(at(2), NONE, SCOPE), undefined);
  // A supporting book that is not this workbook: its sheets cannot be named from here.
  assert.equal(decodeFormula(at(3), NONE, SCOPE), undefined);
});

test('a reference through an unrecognised externals table does not resolve', () => {
  // With a supporting book this reader could not account for, the whole table's indices are
  // untrustworthy — naming a sheet from them could name the wrong one.
  const opaque: FormulaScope = {...SCOPE, selfSupBook: undefined};
  assert.equal(decodeFormula(bytes(0x5a, 0, 0, 0, 0, 0, 0, 0x00, 0xc0), NONE, opaque), undefined);
});

test('a name is cited by its position in the file, counting the ones the model drops', () => {
  assert.equal(decodeFormula(bytes(0x43, 1, 0, 0, 0), NONE, SCOPE), 'Rate');
  assert.equal(decodeFormula(bytes(0x43, 2, 0, 0, 0), NONE, SCOPE), undefined);
});

test('an unrecognised token abandons the formula rather than guessing past it', () => {
  // 0x18 opens the extended-token family (structured references), which this reader does not decode.
  // Its length is not knowable from the ptg alone, so continuing would desynchronise the walk.
  assert.equal(decodeFormula(bytes(0x1e, 1, 0, 0x18, 0x19, 0, 0), NONE, SCOPE), undefined);
});

test('a stream that does not reduce to one value is rejected', () => {
  // Two operands and no operator: a well-formed formula leaves exactly one string behind.
  assert.equal(decodeFormula(bytes(0x1e, 1, 0, 0x1e, 2, 0), NONE, SCOPE), undefined);
  // An operator with nothing to consume.
  assert.equal(decodeFormula(bytes(0x03), NONE, SCOPE), undefined);
});

test('a token running past the end of its own stream is a parse error', () => {
  assert.throws(() => decodeFormula(bytes(0x1e, 1), NONE, SCOPE), XlsbParseError);
});

test('an array constant declaring more elements than any workbook holds is refused', () => {
  // The elements are read from the trailing block, which cannot outrun its record — but the row and
  // column counts are multiplied before any of it is read, so a forged pair must not buy a loop.
  const rgce = bytes(0x60, ...new Array<number>(14).fill(0));
  const rgcb = bytes(0xff, 0xff, 0xff, 0x0f, 0xff, 0xff, 0xff, 0x0f);
  assert.equal(decodeFormula(rgce, rgcb, SCOPE), undefined);
});

test('a deferred formula names the cell it defers to, whose column lives in the extra block', () => {
  assert.deepEqual(formulaAnchor(bytes(0x01, 3, 0, 0, 0), bytes(5, 0, 0, 0)), {row: 3, column: 5});
  // An ordinary formula defers to nothing.
  assert.equal(formulaAnchor(ref(0, 0), NONE), undefined);
  // A deferral whose column is missing is not one.
  assert.equal(formulaAnchor(bytes(0x01, 3, 0, 0, 0), NONE), undefined);
});

test('the function table matches the specification at every run boundary', () => {
  // The table is transcribed as contiguous runs, so a single miscounted entry would shift every name
  // after it. These anchors are the first and last of each run, plus the gaps between them.
  assert.equal(functionNameFor(0x0000), 'COUNT');
  assert.equal(functionNameFor(0x0004), 'SUM');
  assert.equal(functionNameFor(0x0013), 'PI');
  assert.equal(functionNameFor(0x00a9), 'COUNTA');
  assert.equal(functionNameFor(0x00c9), 'UNREGISTER');
  assert.equal(functionNameFor(0x00cc), 'USDOLLAR');
  assert.equal(functionNameFor(0x00d8), 'RANK');
  assert.equal(functionNameFor(0x00db), 'ADDRESS');
  assert.equal(functionNameFor(0x00f8), 'PAUSE');
  assert.equal(functionNameFor(0x00fb), 'RESUME');
  assert.equal(functionNameFor(0x00ff), 'User Defined Function');
  assert.equal(functionNameFor(0x014c), 'TINV');
  assert.equal(functionNameFor(0x014e), 'MOVIE.COMMAND');
  assert.equal(functionNameFor(0x017b), 'RTD');

  for (const gap of [0x00ca, 0x00cb, 0x00d9, 0x00da, 0x00f9, 0x00fa, 0x014d, 0x017c]) {
    assert.equal(functionNameFor(gap), undefined);
  }
});

test('a fixed-arity function knows how many operands belong to its call', () => {
  assert.equal(fixedArityFor('PI'), 0);
  assert.equal(fixedArityFor('ABS'), 1);
  assert.equal(fixedArityFor('ROUND'), 2);
  assert.equal(fixedArityFor('MID'), 3);
  assert.equal(fixedArityFor('REPLACE'), 4);
  // A variadic function has no fixed arity — its call token carries its own count instead.
  assert.equal(fixedArityFor('SUM'), undefined);
  assert.equal(fixedArityFor('IF'), undefined);
});

test('a fixed-arity call with no arity to go on is not decoded', () => {
  // PtgFunc citing SUM: variadic, so the stream states no argument count and the call cannot be
  // reconstructed. Excel only ever writes fixed-arity functions through this token.
  assert.equal(decodeFormula(bytes(...ref(0, 0), 0x41, 0x04, 0x00), NONE, SCOPE), undefined);
  // PtgFunc citing ABS, which takes exactly one.
  assert.equal(decodeFormula(bytes(...ref(0, 0), 0x41, 0x18, 0x00), NONE, SCOPE), 'ABS(A1)');
});
