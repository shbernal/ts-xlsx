// Cluster: tables
//
// Real-world scenario: a worksheet has a table and an anchored image, and the user inserts a row
// above them (a common template-fill pattern). Every table cell-range reference and every image
// anchor must shift to stay pinned to the same logical data. When they do not, the produced file has
// table ranges and image anchors pointing at the wrong cells, which Excel reports as corrupt. The
// row splice moves the cell data but strands the table range and image anchor at their old
// coordinates. (The companion merge-shift on splice is locked by `splice-rows-preserves-merged-cells`;
// this case covers table ranges and image anchors, plus table column-name uniqueness.)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'splice-rows-updates-table-and-image-refs',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Inserting a row above a table and an anchored image shifts both the table’s cell range and the ' +
    'image’s anchor down by the inserted rows, and authoring a table with duplicate column names is ' +
    'rejected — so a template fill does not strand refs at stale coordinates or emit an invalid table.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'inserting a row above a table shifts the table’s cell range down',
      baseline: 'fail',
      async expect(api, assert) {
        const {tableRef} = await api.spliceShiftsRefs();
        assert.strictEqual(tableRef, 'A4:B6', `the table at A3:B5 must shift to A4:B6 after inserting a top row; got ${JSON.stringify(tableRef)}`);
      },
    },
    {
      name: 'inserting a row above an anchored image shifts the image’s from-anchor down',
      baseline: 'fail',
      async expect(api, assert) {
        const {imageFromRow} = await api.spliceShiftsRefs();
        assert.strictEqual(imageFromRow, 6, `the image anchored from row 5 must shift to row 6; got ${JSON.stringify(imageFromRow)}`);
      },
    },
    {
      name: 'authoring a table with duplicate column names is rejected',
      baseline: 'fail',
      async expect(api, assert) {
        const {dupColumnNamesRejected} = await api.spliceShiftsRefs();
        assert.strictEqual(dupColumnNamesRejected, true, 'duplicate table column names must be rejected, not written into an invalid table');
      },
    },
  ],
};
