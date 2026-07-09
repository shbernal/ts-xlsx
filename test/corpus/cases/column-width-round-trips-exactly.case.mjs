// Cluster: core-model
//
// Real-world scenario: a user sets a column width to a round number (12) expecting
// that exact stored value to survive a write→read cycle. A recurring confusion is
// that Excel *displays* a slightly different figure (a character-width vs. pixel
// padding offset of ~0.71), leading people to believe the library mangles the
// value. The stored model value must round-trip byte-for-byte — no silent unit
// drift added or subtracted on write or read.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void> }} Behavior */

const WHOLE = {sheets: [{name: 'S', columns: [{index: 1, width: 12}]}]};
const FRACTION = {sheets: [{name: 'S', columns: [{index: 1, width: 8.43}]}]};

export default {
  id: 'column-width-round-trips-exactly',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 744},
  cluster: 'core-model',
  description:
    'A column width set in the model survives a write→read round-trip as the exact ' +
    'same value, with no unit conversion silently added (the ~0.71 users see is ' +
    "Excel's display padding, not a stored-value change).",

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a whole-number width reads back unchanged',
      baseline: 'pass',
      async expect(api, assert) {
        const {width} = (await api.roundtripWorkbook(WHOLE)).sheets.S.columns[1];
        assert.strictEqual(width, 12, `expected width 12, got ${width}`);
      },
    },
    {
      name: 'a fractional width reads back unchanged',
      baseline: 'pass',
      async expect(api, assert) {
        const {width} = (await api.roundtripWorkbook(FRACTION)).sheets.S.columns[1];
        assert.strictEqual(width, 8.43, `expected width 8.43, got ${width}`);
      },
    },
  ],
};
