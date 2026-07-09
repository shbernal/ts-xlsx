// Cluster: defined-names
//
// Real-world scenario: a workbook has two defined names that share the same name but are scoped
// to different sheets (each `<definedName>` carries a distinct `localSheetId`) — a common Excel
// pattern where "namedrange" means one range on Page1 and a different range on Page2. Today the
// reader keys defined names by name alone, so the two collide on load and only the last one
// survives; the first sheet's scoped range is lost. Same-named, differently-scoped defined names
// must both be retained, each associated with its owning sheet.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'same-named-defined-names-scoped-per-sheet/sample.xlsx';

export default {
  id: 'same-named-defined-names-scoped-per-sheet',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1497},
  cluster: 'defined-names',
  description:
    'Two defined names sharing a name but scoped to different sheets are both retained on load ' +
    '— each keeps its own sheet-scoped range instead of colliding so that only the last survives.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'both sheet-scoped ranges of a same-named defined name are retained',
      baseline: 'fail',
      async expect(api, assert) {
        const {names} = await api.readFixtureDefinedNames(FIXTURE);
        const ranges = names.namedrange || [];
        assert.ok(
          ranges.includes('Page1!$A$1:$B$1'),
          `the Page1-scoped range must survive the load; got ${JSON.stringify(ranges)}`
        );
        assert.ok(
          ranges.includes('Page2!$A$1:$B$1'),
          `the Page2-scoped range must survive the load; got ${JSON.stringify(ranges)}`
        );
      },
    },
  ],
};
