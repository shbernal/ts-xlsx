// Cluster: styles
//
// Real-world scenario: a package's worksheet declares a column-level style — a `<col>` element
// carrying a styleId (from a column number-format, font, width-with-style, etc.) — but the package
// has no styles part for that id to resolve against. This happens with partially-assembled packages,
// producers that emit column styles without a complete styles.xml, or files where the styles part was
// stripped. Reconciling the column looks the styleId up in the styles collection; when that
// collection is absent the reader dereferences `undefined.getStyleModel(...)` and aborts the entire
// load with an internal TypeError. A dangling column-style reference is malformed input the reader
// must tolerate — loading the worksheet without the unresolved style — not a crash.
//
// The fixture is a normal single-column-styled workbook whose styles part has been removed while the
// worksheet's `<col>` styleId reference remains, reproducing the dangling reference.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'column-style-without-styles-part/sample.xlsx';

export default {
  id: 'column-style-without-styles-part-must-not-crash-reader',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A workbook whose worksheet has a column with a style reference but no styles part loads without ' +
    'throwing — the reader tolerates the unresolvable column styleId instead of dereferencing an ' +
    'undefined styles collection, and the worksheet is still available.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'loading a workbook whose column style has no styles part does not throw',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(ok, true, `the load must not throw on a dangling column-style reference; got ${error}`);
      },
    },
    {
      name: 'the worksheet is still present after tolerating the unresolved column style',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(
          sheetNames && sheetNames.includes('Sheet1'),
          `the worksheet must load despite the missing styles part; got ${JSON.stringify(sheetNames)}`
        );
      },
    },
  ],
};
