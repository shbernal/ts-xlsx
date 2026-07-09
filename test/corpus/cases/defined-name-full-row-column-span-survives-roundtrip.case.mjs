// Cluster: tables
//
// Real-world scenario: a workbook's Name Manager holds a defined name whose reference spans
// entire rows or entire columns rather than a bounded rectangle — e.g. a name pointing at
// "Sheet2!$1:$5" (rows 1–5, every column) or the whole-axis forms "$A:$C" (all rows of three
// columns) and "$1:$1048576" (all columns of a row block). Excel creates these routinely.
// Today an over-strict range-address check that demands explicit column *and* row bounds
// silently discards any open-ended span: `definedNames.add` accepts the reference without
// error yet the name never lands in the model, so it is absent from the written file — and a
// file that already declares such a name reads back with the name gone. Bounded references
// (both corners fully qualified) survive; only the open-ended spans vanish. Full-row and
// full-column named ranges must survive read and write exactly like bounded ones.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'defined-name-full-row-column-span-survives-roundtrip/whole-row-named-range.xlsx';

// A control (bounded reference) alongside the open-ended spans under test, so a baseline flip
// is unambiguous: the bug is specifically about references lacking explicit row/column bounds.
const BOUNDED = {definedNames: [{name: 'Bounded', ranges: ['Sheet1!$A$1:$C$5']}], sheets: [{name: 'Sheet1'}, {name: 'Sheet2'}]};
const FULL_ROW = {definedNames: [{name: 'FullRow', ranges: ['Sheet2!$1:$5']}], sheets: [{name: 'Sheet1'}, {name: 'Sheet2'}]};
const FULL_COL = {definedNames: [{name: 'FullCol', ranges: ['Sheet1!$A:$C']}], sheets: [{name: 'Sheet1'}, {name: 'Sheet2'}]};

export default {
  id: 'defined-name-full-row-column-span-survives-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1174},
  cluster: 'tables',
  description:
    'A defined name whose reference is an entire-row or entire-column span survives a ' +
    'read/write round-trip instead of being silently dropped by over-strict address ' +
    'validation — bounded references already survive; the open-ended spans must too.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a bounded-reference defined name survives a round-trip (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {definedNames} = await api.roundtripWorkbook(BOUNDED);
        assert.deepStrictEqual(definedNames.Bounded, ['Sheet1!$A$1:$C$5'], 'a fully-bounded name round-trips');
      },
    },
    {
      name: 'a full-row-span defined name survives a write→read round-trip',
      baseline: 'fail',
      async expect(api, assert) {
        const {definedNames} = await api.roundtripWorkbook(FULL_ROW);
        assert.deepStrictEqual(
          definedNames.FullRow,
          ['Sheet2!$1:$5'],
          `a whole-row span must not be dropped; got ${JSON.stringify(definedNames)}`
        );
      },
    },
    {
      name: 'a full-column-span defined name survives a write→read round-trip',
      baseline: 'fail',
      async expect(api, assert) {
        const {definedNames} = await api.roundtripWorkbook(FULL_COL);
        assert.deepStrictEqual(
          definedNames.FullCol,
          ['Sheet1!$A:$C'],
          `a whole-column span must not be dropped; got ${JSON.stringify(definedNames)}`
        );
      },
    },
    {
      name: 'a full-row-span defined name declared by a real file is read back',
      baseline: 'fail',
      async expect(api, assert) {
        const {names} = await api.readFixtureDefinedNames(FIXTURE);
        assert.deepStrictEqual(
          names.RangeTest,
          ['Sheet2!$1:$5'],
          `the file declares RangeTest=Sheet2!$1:$5; reader must expose it, got ${JSON.stringify(names)}`
        );
      },
    },
  ],
};
