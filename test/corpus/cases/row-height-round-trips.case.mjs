// Cluster: core-model
//
// Real-world scenario: a report sets an explicit row height (to make a header band
// tall, or to give wrapped text room). On reopening the file the rows must keep
// that height. A notorious failure mode is heights silently not persisting, so the
// rows reopen collapsed and the data looks "gone" until each row is dragged open.
// An explicitly set height must survive a write→read round-trip exactly.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void> }} Behavior */

const TALL = {sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x'}], rows: [{index: 1, height: 300}]}]};
const MODEST = {sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x'}], rows: [{index: 1, height: 15}]}]};

export default {
  id: 'row-height-round-trips',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 653},
  cluster: 'core-model',
  description:
    'An explicit row height set in the model survives a write→read round-trip as the ' +
    'same value, so reopened rows keep their intended height instead of collapsing.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a large explicit height survives the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {height} = (await api.roundtripWorkbook(TALL)).sheets.S.rows[1];
        assert.strictEqual(height, 300, `expected height 300, got ${height}`);
      },
    },
    {
      name: 'a modest explicit height survives the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {height} = (await api.roundtripWorkbook(MODEST)).sheets.S.rows[1];
        assert.strictEqual(height, 15, `expected height 15, got ${height}`);
      },
    },
  ],
};
