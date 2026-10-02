// Cluster: formulas
//
// Real-world scenario: a sheet uses columns whose letters spell a function, such as T, N or PI, and its
// formulas reference them absolutely (`$T$1`, `T$1`), as whole columns (`T:T`) or inside a LET or
// LAMBDA that binds a parameter of the same name. The function-value prefix `_xleta.` and the parameter
// prefix `_xlpm.` belong to names, and a column of a reference is neither: `$_xleta.T$1` or
// `$_xlpm.T$1` names no cell, and Excel refused to open a workbook that stored it.
//
// The fixture was saved by Excel Desktop 16.0 (build 20326) from one sheet S1 holding 5 in T1, 7 in N2
// and 1, 2 and 3 in PI1:PI3, with the formulas below typed into column B through `Range.Formula2`.
// `test/corpus/fixtures/excel-oracle/references-spelled-like-functions.json` records the values Excel
// computed for them, the same values for this library's own write of them, and the refusal of the
// write made before the fix.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'reference-spelled-like-a-function/excel-saved.xlsx';

// `read` is the formula text Excel's `Range.Formula` reports, without its `=`; `stored` is the text
// Excel stored in the package.
const EXCEL: Record<string, {read: string; stored: string}> = {
  'S1!B1': {read: 'IF($T$1>0,1,0)', stored: 'IF($T$1>0,1,0)'},
  'S1!B2': {read: 'T$1+$N$2', stored: 'T$1+$N$2'},
  'S1!B3': {read: 'SUM($PI$1:$PI$3)', stored: 'SUM($PI$1:$PI$3)'},
  'S1!B4': {read: 'SUM(T:T)', stored: 'SUM(T:T)'},
  'S1!B5': {read: 'SUM(N:N)+SUM($PI:$PI)', stored: 'SUM(N:N)+SUM($PI:$PI)'},
  'S1!B6': {read: "SUM('S1'!T:T)", stored: "SUM('S1'!T:T)"},
  'S1!B7': {read: 'LET(T,1,$T$1+T)', stored: '_xlfn.LET(_xlpm.T,1,$T$1+_xlpm.T)'},
  'S1!B8': {read: 'LET(T,1,T$1+$T1+T)', stored: '_xlfn.LET(_xlpm.T,1,T$1+$T1+_xlpm.T)'},
  'S1!B9': {read: 'LET(N,2,SUM($N:$N)+N)', stored: '_xlfn.LET(_xlpm.N,2,SUM($N:$N)+_xlpm.N)'},
  'S1!B10': {read: 'LAMBDA(T,$T$1+T)(1)', stored: '_xlfn.LAMBDA(_xlpm.T,$T$1+_xlpm.T)(1)'},
  // Without a `$`, the parameter wins: Excel reads both ends of `T:T` as the bound `T`.
  'S1!B11': {
    read: 'LET(T,1,SUM(T:T)+T)',
    stored: '_xlfn.LET(_xlpm.T,1,SUM(_xlpm.T:_xlpm.T)+_xlpm.T)',
  },
};

const project = (field: 'read' | 'stored') =>
  Object.fromEntries(Object.entries(EXCEL).map(([key, spelling]) => [key, spelling[field]]));

export default {
  id: 'a-reference-spelled-like-a-function-stays-a-reference',
  cluster: 'formulas',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  description:
    'A cell or column reference whose column letters spell a function, or a LET or LAMBDA parameter, ' +
    'is stored as the reference it is, without the _xleta. or _xlpm. prefix a name of that spelling ' +
    'would take, and a workbook Excel saved writes back in the spelling Excel stored.',
  behavior: [
    {
      name: 'reading a workbook Excel saved gives every formula the text Excel reports for it',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.fixtureFormulaSpellings(FIXTURE).read, project('read'));
      },
    },
    {
      name: 'writing that workbook back stores every formula in the spelling Excel stored',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.fixtureFormulaSpellings(FIXTURE).written, project('stored'));
      },
    },
    {
      name: 'an authored reference whose column spells a function or a parameter is stored as typed',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(
          api.storedFormulas({
            'S1!B1': 'IF($T$1>0,1,0)',
            'S1!B2': 'SUM(T:T)+N$2',
            'S1!B3': 'LET(T,1,$T$1+T)',
            'S1!B4': 'LET(T,1,SUM(T:T)+T)',
          }),
          {
            'S1!B1': 'IF($T$1>0,1,0)',
            'S1!B2': 'SUM(T:T)+N$2',
            'S1!B3': '_xlfn.LET(_xlpm.T,1,$T$1+_xlpm.T)',
            'S1!B4': '_xlfn.LET(_xlpm.T,1,SUM(_xlpm.T:_xlpm.T)+_xlpm.T)',
          },
        );
      },
    },
  ],
} satisfies Case;
