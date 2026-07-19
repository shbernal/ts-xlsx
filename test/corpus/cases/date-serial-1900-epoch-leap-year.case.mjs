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
      baseline: 'pass',
      async expect(api, assert) {
        const {A1} = await api.readFixtureCells(FIXTURE, ['A1']);
        assert.strictEqual(A1.type, 'date', 'the cell is a date');
        assert.strictEqual(A1.value.date, '1900-01-01T00:00:00.000Z', `serial 1 must be 1900-01-01; got ${JSON.stringify(A1.value)}`);
      },
    },
    {
      name: 'serial 2 reads as 1900-01-02 (consecutive days)',
      baseline: 'pass',
      async expect(api, assert) {
        const {A2} = await api.readFixtureCells(FIXTURE, ['A2']);
        assert.strictEqual(A2.value.date, '1900-01-02T00:00:00.000Z', `serial 2 must be 1900-01-02; got ${JSON.stringify(A2.value)}`);
      },
    },
    {
      // The boundary case: serial 60 is the phantom 1900-02-29 that never existed. Every serial at or
      // below 59 is therefore one day later than a naive "days since 1899-12-31" offset would place
      // it — serial 59 is 1900-02-28, the real day just before the phantom. A reader that ignores the
      // phantom day reads serial 59 as 1900-02-27, one day early.
      name: 'serial 59 reads as 1900-02-28 (the real day just below the phantom 1900-02-29)',
      baseline: 'pass',
      async expect(api, assert) {
        const {A59} = await api.readFixtureCells(FIXTURE, ['A59']);
        assert.strictEqual(A59.value.date, '1900-02-28T00:00:00.000Z', `serial 59 must be 1900-02-28; got ${JSON.stringify(A59.value)}`);
      },
    },
    {
      name: 'serial 61 reads as 1900-03-01 (the day after the phantom leap day)',
      baseline: 'pass',
      async expect(api, assert) {
        const {A61} = await api.readFixtureCells(FIXTURE, ['A61']);
        assert.strictEqual(A61.value.date, '1900-03-01T00:00:00.000Z', `serial 61 must be 1900-03-01; got ${JSON.stringify(A61.value)}`);
      },
    },
  ],
};
