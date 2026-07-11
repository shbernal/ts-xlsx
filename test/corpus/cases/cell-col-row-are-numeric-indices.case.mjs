// Cluster: types
//
// Real-world scenario: each cell exposes its position through a column index and a row index —
// 1-based integers describing where the cell lives. Consumers rely on them as numbers for arithmetic
// (comparing positions, computing offsets, indexing parallel arrays). The runtime has always returned
// numbers, but the legacy published TypeScript declared the accessors as `string`, so strict code
// treating them as numbers failed to type-check against a runtime that was already numeric. This
// locks the runtime contract (numeric 1-based indices) the type surface must match; the declaration
// honesty itself is tracked in `cell-full-address-descriptor-numeric-row-col`.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'cell-col-row-are-numeric-indices',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    "A cell's col and row accessors are 1-based numbers at runtime, matching its actual position — " +
    'not strings.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: "a cell's col and row are of runtime type number",
      baseline: 'pass',
      async expect(api, assert) {
        const {colType, rowType} = await api.cellColRowTypes('B3');
        assert.strictEqual(colType, 'number', 'col is a number at runtime');
        assert.strictEqual(rowType, 'number', 'row is a number at runtime');
      },
    },
    {
      name: 'col and row are the 1-based indices of the cell position (B3 → col 2, row 3)',
      baseline: 'pass',
      async expect(api, assert) {
        const {col, row} = await api.cellColRowTypes('B3');
        assert.strictEqual(col, 2, 'column B is index 2');
        assert.strictEqual(row, 3, 'row 3 is index 3');
      },
    },
  ],
};
