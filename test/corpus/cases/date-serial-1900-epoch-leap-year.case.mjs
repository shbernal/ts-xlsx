// Cluster: types
//
// Real-world scenario: date-formatted cells store a numeric serial in the 1900 date system. Excel
// treats serial 1 as 1900-01-01 and — because of a deliberately-reproduced 1900 leap-year quirk
// (serial 60 is the non-existent 1900-02-29) — serials 1..59 map to 1900-01-01..1900-02-28. A
// reader that uses a straight 1899-12-30 epoch without accounting for the phantom day reads every
// pre-March-1900 serial one day early (serial 1 → 1899-12-31 instead of 1900-01-01), so imported
// dates are wrong. A date-formatted serial must read as the calendar date Excel displays for it.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// The fixture's column A holds serials 1,2,3,… as date-formatted cells.
const FIXTURE = 'date-serial-1900-epoch-leap-year/serials.xlsx';

export default {
  id: 'date-serial-1900-epoch-leap-year',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1928},
  cluster: 'types',
  description:
    'A date-formatted serial in the 1900 date system reads as the calendar date Excel displays: ' +
    'serial 1 is 1900-01-01 (not 1899-12-31), and consecutive serials map to consecutive days, ' +
    'accounting for the 1900 leap-year quirk.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'serial 1 reads as 1900-01-01',
      baseline: 'fail',
      async expect(api, assert) {
        const {A1} = await api.readFixtureCells(FIXTURE, ['A1']);
        assert.strictEqual(A1.type, 'date', 'the cell is a date');
        assert.strictEqual(A1.value.date, '1900-01-01T00:00:00.000Z', `serial 1 must be 1900-01-01; got ${JSON.stringify(A1.value)}`);
      },
    },
    {
      name: 'serial 2 reads as 1900-01-02 (consecutive days)',
      baseline: 'fail',
      async expect(api, assert) {
        const {A2} = await api.readFixtureCells(FIXTURE, ['A2']);
        assert.strictEqual(A2.value.date, '1900-01-02T00:00:00.000Z', `serial 2 must be 1900-01-02; got ${JSON.stringify(A2.value)}`);
      },
    },
  ],
};
