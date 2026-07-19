// Cluster: images
//
// Real-world scenario: a worksheet declares a drawing relationship (a <drawing r:id> in the sheet,
// with a matching entry in the worksheet's rels) but the drawing part it points to does not resolve —
// it is absent or unreadable in the package. This happens with files from some producers, or when a
// package is partially assembled. Reading such a workbook must not abort the whole load with an
// internal null-dereference (the reader looking up an undefined drawing and reading `.anchors` off
// it); the worksheet should load, simply without the unresolved drawing's images.
//
// The fixture is a normal image-bearing workbook whose single drawing part has been removed while its
// worksheet drawing relationship and <drawing> element remain — reproducing the dangling reference.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'load-worksheet-missing-drawing-part/sample.xlsx';

export default {
  id: 'load-worksheet-missing-drawing-part-does-not-crash',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'A workbook whose worksheet references a drawing part that does not resolve loads without ' +
    'throwing — the reader tolerates the dangling drawing reference instead of dereferencing ' +
    'undefined, and the worksheet is still available.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'loading a workbook with an unresolved drawing reference does not throw',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(ok, true, `the load must not throw on a dangling drawing reference; got ${error}`);
      },
    },
    {
      name: 'the worksheet is still present after tolerating the dangling drawing',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(sheetNames && sheetNames.includes('Sheet1'), `the worksheet must load; got ${JSON.stringify(sheetNames)}`);
      },
    },
  ],
};
