// Cluster: images
//
// Real-world scenario: a user opens an .xlsx that contains a chart — a worksheet drawing carrying a
// chart graphicFrame rather than a picture/image. They only want to read data cells. During load the
// reader reconciles worksheet drawings against their relationship parts; a drawing that holds a chart
// (and therefore no picture anchors) must not make the reconciliation dereference an undefined
// drawing model and throw. Excel opens such a file fine. The reader must tolerate chart-bearing
// (non-picture) drawings: the load succeeds, sheet data is readable, and the chart drawing is simply
// skipped for image purposes rather than aborting the whole workbook.
//
// The fixture is a workbook whose worksheet drawing was replaced with a chart graphicFrame anchor
// (no <xdr:pic>), reproducing the "drawing with no picture anchors" shape without needing a full
// application-authored chart.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'load-workbook-with-chart-drawing/sample.xlsx';

export default {
  id: 'load-workbook-with-chart-drawing-does-not-crash',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'Loading a workbook whose worksheet drawing carries a chart graphicFrame (no picture anchors) ' +
    'completes without throwing and the worksheets and their cell values are readable, rather than ' +
    'the reconciliation crashing on a drawing with no anchors.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a workbook whose drawing is a chart graphicFrame loads without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `a chart-bearing drawing must be tolerated; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'the worksheets are recovered after loading the chart-bearing workbook',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.deepStrictEqual(
          sheetNames,
          ['Sheet1', 'Second'],
          'both worksheets survive the tolerant load',
        );
      },
    },
    {
      name: 'cell values on the chart’s sheet are readable',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = await api.readFixtureCells(FIXTURE, ['A1', 'B1']);
        assert.strictEqual(cells.A1.value, 'data', 'a text cell on the chart sheet reads back');
        assert.strictEqual(cells.B1.value, 7, 'a numeric cell on the chart sheet reads back');
      },
    },
  ],
};
