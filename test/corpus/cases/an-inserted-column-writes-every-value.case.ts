// Cluster: address-decoding
//
// Real-world scenario: a sheet is being built up, and a whole column of values is inserted at a
// position -- an id column in front of a report, a computed column between two existing ones. A
// column's values are indexed by row, so the array names rows the sheet mostly does not hold yet:
// on a sheet being filled in, that is nearly all of them. The insert wrote its values while walking
// the rows the grid already had, so a value whose row did not exist had nowhere to land and was
// simply dropped. Inserting into an empty sheet wrote nothing whatsoever and reported no error;
// inserting beside a single cell wrote the first value and lost the rest. `addColumn`, given the
// same array, materialised every row of it, so the two column-writing entry points disagreed about
// their own argument with nothing to tell the caller which one was right.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'an-inserted-column-writes-every-value',
  provenance: {source: 'grid-edit-audit'},
  cluster: 'address-decoding',
  description:
    "An inserted column's values are indexed by row and land on every row they name, including " +
    'rows the sheet does not hold yet. Inserting a column and appending one place the same array ' +
    'identically; a hole in it leaves its row untouched on both paths.',

  behavior: [
    {
      name: 'inserting a column into an empty sheet writes the whole column',
      async expect(api: CorpusApi, assert: Assert) {
        const {intoEmpty} = await api.columnInsertMaterialisesEveryValue();
        assert.deepStrictEqual(
          intoEmpty.column,
          ['x', 'y', 'z'],
          'an empty sheet held no row to write onto, so this used to write nothing at all',
        );
        assert.strictEqual(intoEmpty.rowCount, 3, 'the sheet grew to hold what it was given');
      },
    },
    {
      name: 'the values past the last existing row are not dropped',
      async expect(api: CorpusApi, assert: Assert) {
        const {intoShortGrid} = await api.columnInsertMaterialisesEveryValue();
        assert.deepStrictEqual(intoShortGrid.column, ['x', 'y', 'z']);
        assert.strictEqual(
          intoShortGrid.shifted,
          'was here',
          'the existing cell still shifted right',
        );
      },
    },
    {
      name: 'inserting a column and appending one place the same array the same way',
      async expect(api: CorpusApi, assert: Assert) {
        const {intoEmpty, appended} = await api.columnInsertMaterialisesEveryValue();
        assert.deepStrictEqual(
          intoEmpty.column,
          appended.column,
          'the two column-writing entry points read their argument identically',
        );
        assert.strictEqual(intoEmpty.rowCount, appended.rowCount);
      },
    },
  ],
} satisfies Case;
