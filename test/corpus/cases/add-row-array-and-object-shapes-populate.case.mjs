// Cluster: tables
//
// Real-world scenario: a user with keyed columns appends rows in three interchangeable shapes — a
// dense positional array (values map left-to-right onto columns), a sparse 1-based array (indices
// place values into specific columns, gaps stay empty), and a key/value object keyed by the column
// keys. All three must land their data, and a batch append mixing array- and object-shaped rows must
// populate every row, not only the object-keyed ones. Types (number, date) written via array append
// survive a save/reload.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'add-row-array-and-object-shapes-populate',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Appending rows as a dense positional array, a sparse 1-based array, and a key/value object each ' +
    'places data in the right columns; a mixed batch populates every row (not just object-keyed ones), ' +
    'and numeric/date types written by array append survive a round-trip.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a dense positional array places elements into columns left-to-right',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.appendRowShapes();
        assert.deepStrictEqual(
          [rows[2].A, rows[2].B, rows[2].C],
          ['a', 'b', 'c'],
          'a dense array maps first→A, second→B, third→C',
        );
      },
    },
    {
      name: 'a sparse 1-based array places values at their indices and leaves gaps empty',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.appendRowShapes();
        assert.strictEqual(rows[3].A, 'x', 'index 1 lands in column A');
        assert.strictEqual(rows[3].B, null, 'the gap at index 2 stays empty');
        assert.strictEqual(rows[3].C, 'z', 'index 3 lands in column C');
      },
    },
    {
      name: 'a key/value object places values under their matching column keys',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.appendRowShapes();
        assert.deepStrictEqual(
          [rows[4].A, rows[4].B],
          ['o1', 'o2'],
          'object values land under k1/k2',
        );
      },
    },
    {
      name: 'a mixed batch of array- and object-shaped rows populates every row',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.appendRowShapes();
        assert.deepStrictEqual(
          [rows[6].A, rows[6].B],
          ['m1', 'm2'],
          'the array-shaped batch row is populated',
        );
        assert.strictEqual(rows[7].A, 'n1', 'the object-shaped batch row is populated');
      },
    },
    {
      name: 'numeric and date types written by array append survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.appendRowShapes();
        assert.strictEqual(rows[5].A, 7, 'the number stays numeric, not stringified');
        assert.deepStrictEqual(
          rows[5].B,
          {date: '2021-01-02T00:00:00.000Z'},
          'the date stays a date',
        );
      },
    },
    {
      // Array detection must be structural, not realm-bound: an array built in a foreign realm (a
      // Node vm context, or a browser iframe) is still an array — Array.isArray says so — but code
      // that checks `instanceof Array` or the constructor identity mistakes it for a plain object
      // and populates no cells. Feeding such a row must place one value per element like any array.
      name: 'a row from an array built in another realm still populates one cell per element',
      baseline: 'pass',
      async expect(api, assert) {
        const {isArrayCrossRealm, a, b, c} = await api.crossRealmArrayRow();
        assert.strictEqual(
          isArrayCrossRealm,
          true,
          'Array.isArray recognizes the cross-realm array (oracle)',
        );
        assert.deepStrictEqual(
          [a, b, c],
          [10, 20, 30],
          'the cross-realm array must fill columns A/B/C, not leave them null',
        );
      },
    },
  ],
};
