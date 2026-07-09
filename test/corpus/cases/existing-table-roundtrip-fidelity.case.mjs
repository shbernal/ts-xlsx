// Cluster: tables
//
// Real-world scenario: a workbook contains a defined table (an Excel Table / ListObject) over a region
// of a sheet. Loading that workbook and writing it back out must preserve the table definition so the
// result opens cleanly without a repair prompt: the table's reference range must survive in the written
// table part rather than being dropped or shifted, the part itself must not vanish, and degenerate
// shapes — a table with an empty body or a single data row — must round-trip without error or injected
// padding rows.
//
// (A related hazard — a table whose HEADER ROW is hidden must stay valid on save — needs an
// Excel-authored fixture that declares a hidden header, so it is recorded in the clone/merge and
// table spec notes rather than asserted here; this case covers the fidelity that a spec-built table
// can exercise.)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const table = (rows, ref) => ({
  sheets: [{name: 'S', tables: [{name: 'T', ref, headers: ['C1', 'C2'], rows}]}],
});

export default {
  id: 'existing-table-roundtrip-fidelity',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A defined table survives a load→save round-trip with its reference range intact and its table ' +
    'part not dropped, including the degenerate empty-body and single-data-row shapes that must ' +
    'round-trip without error.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: "a table's reference range is written and survives a load→save round-trip",
      baseline: 'pass',
      async expect(api, assert) {
        const {write, roundtrip, loadOk} = await api.roundtripSpecTableFacts(table([[1, 2], [3, 4]], 'A1:B3'));
        assert.ok(loadOk, 'the written table loads without throwing');
        assert.strictEqual(write[0].ref, 'A1:B3', 'the ref is written as authored');
        assert.strictEqual(roundtrip[0].ref, write[0].ref, 'the ref is unchanged after the round-trip');
      },
    },
    {
      name: 'the table part is not dropped across a load→save round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {roundtrip} = await api.roundtripSpecTableFacts(table([[1, 2], [3, 4]], 'A1:B3'));
        assert.strictEqual(roundtrip.length, 1, 'the table part survives the round-trip');
        assert.ok(roundtrip[0].wellFormed, 'the round-tripped table XML is well-formed');
      },
    },
    {
      name: 'an empty-body table round-trips without error and keeps its part',
      baseline: 'pass',
      async expect(api, assert) {
        const {roundtrip, loadOk, loadError} = await api.roundtripSpecTableFacts(table([], 'A1:B1'));
        assert.ok(loadOk, `an empty-body table must load without throwing; got ${loadError}`);
        assert.strictEqual(roundtrip.length, 1, 'the empty-body table part survives');
      },
    },
    {
      name: 'a single-data-row table round-trips without error',
      baseline: 'pass',
      async expect(api, assert) {
        const {roundtrip, loadOk, loadError} = await api.roundtripSpecTableFacts(table([[1, 2]], 'A1:B2'));
        assert.ok(loadOk, `a single-row table must load without throwing; got ${loadError}`);
        assert.strictEqual(roundtrip.length, 1, 'the single-row table part survives');
      },
    },
  ],
};
