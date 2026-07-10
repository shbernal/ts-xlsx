// Cluster: tables
//
// Real-world scenario: a worksheet has a formatted table, and the author then merges a cell range
// that overlaps the table's region. Excel forbids merged cells inside a formatted table, so the
// resulting package opens as corrupt (Excel repairs it, dropping the merge or the table). The write
// path should treat an overlapping merge as a conflict and surface it, rather than silently emitting
// Excel-invalid geometry. A merge that lies entirely outside every table is valid and must keep
// working, and a table with no overlapping merge must round-trip cleanly.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const table = () => ({name: 'T', ref: 'A1', headers: ['H1', 'H2'], rows: [['a', 1], ['b', 2]]});

// The table occupies A1:B3 (header + two data rows). A merge over A2:B2 lands inside it.
const overlappingSpec = {sheets: [{name: 'S', tables: [table()], merges: ['A2:B2']}]};
// A merge well clear of the table region.
const disjointSpec = {sheets: [{name: 'S', cells: [{ref: 'D5', value: 'x'}], tables: [table()], merges: ['D5:E5']}]};
const noMergeSpec = {sheets: [{name: 'S', tables: [table()]}]};

export default {
  id: 'merge-overlapping-a-table-region-is-rejected',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A merged range that overlaps a formatted table’s region is a conflict the write path should ' +
    'surface (Excel forbids merges inside a table and treats the file as corrupt), while a merge ' +
    'entirely outside every table writes cleanly and a table with no overlapping merge round-trips.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'merging a range that overlaps a table region surfaces a conflict instead of writing a corrupt file',
      baseline: 'fail',
      async expect(api, assert) {
        const result = await api.tryWriteWorkbook(overlappingSpec);
        assert.strictEqual(
          result.ok,
          false,
          'a merge overlapping a table must be rejected, not silently written as Excel-invalid geometry'
        );
      },
    },
    {
      name: 'a merge entirely outside every table writes successfully',
      baseline: 'pass',
      async expect(api, assert) {
        const result = await api.tryWriteWorkbook(disjointSpec);
        assert.strictEqual(result.ok, true, `a disjoint merge must write; got ${JSON.stringify(result.error)}`);
      },
    },
    {
      name: 'a table with no overlapping merge round-trips with its geometry intact',
      baseline: 'pass',
      async expect(api, assert) {
        const {tables} = await api.inspectPackage(noMergeSpec);
        assert.strictEqual(tables.length, 1, 'the table part is written');
        assert.strictEqual(tables[0].ref, 'A1:B3', 'the table geometry is intact');
        assert.strictEqual(tables[0].xmlWellFormed, true, 'the table part is well-formed');
      },
    },
  ],
};
