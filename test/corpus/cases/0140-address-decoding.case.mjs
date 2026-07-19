// Provenance: exceljs/exceljs#140 — "col-cache.js: Cannot read property '0' of null"
// Cluster: address-decoding
//
// Real-world .xlsx files declare defined names that reference whole rows or whole
// columns, e.g. `MyWorksheet!$A:$A` (full column) or `'Some Text'!$1:$1` (full row)
// — see the reproduction XML in the issue thread. Decoding these must not crash and
// must not leak the literal strings "undefined"/"NaN" into serialized addresses.
//
// The original TypeError crash was fixed upstream, so the no-throw behaviors are
// GREEN regression locks. But full-row range decoding still emits garbage
// (`$col$row: "$undefined$1"`, `dimensions: "NaN:NaN"`) — captured here as a RED
// live defect the rewrite must fix. `baseline` records the *current legacy* result
// so the runner can tell a known-open bug from a fresh regression.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => void }} Behavior */

export default {
  id: '0140-address-decoding',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 140, url: 'https://github.com/exceljs/exceljs/issues/140'},
  cluster: 'address-decoding',
  description: 'Defined names referencing whole rows/columns must decode without crashing or leaking undefined/NaN.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'decodeAddress("$1") — a full-row absolute reference — does not throw',
      baseline: 'pass',
      expect(api, assert) {
        const addr = api.decodeAddress('$1');
        assert.strictEqual(addr.row, 1, 'row should be 1');
        assert.strictEqual(addr.col, undefined, 'a row-only reference has no column');
      },
    },
    {
      name: 'decodeRange("$1:$1") — a full-row range — resolves its known row bounds',
      baseline: 'pass',
      expect(api, assert) {
        const range = api.decodeRange('$1:$1');
        assert.strictEqual(range.top, 1, 'top row should be 1');
        assert.strictEqual(range.bottom, 1, 'bottom row should be 1');
      },
    },
    {
      name: 'decodeRange("$1:$1") — serialized form leaks no "undefined"/"NaN"',
      baseline: 'pass',
      expect(api, assert) {
        const serialized = JSON.stringify(api.decodeRange('$1:$1'));
        assert.ok(!serialized.includes('undefined'), `serialized range leaks "undefined": ${serialized}`);
        assert.ok(!serialized.includes('NaN'), `serialized range leaks "NaN": ${serialized}`);
      },
    },
  ],
};
