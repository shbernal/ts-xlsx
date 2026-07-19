// Cluster: xlsx-io
//
// Real-world scenario: a cell links to a single-page-app URL whose meaningful part
// lives in the '#' fragment, e.g. http://host/ui/#/case/2007720723. Excel stores a
// hyperlink as a relationship Target plus an optional location; a writer that mixes
// these up can drop everything after the '#'. Writing then reading the workbook must
// return the complete URL, fragment included, and keep the cell's display text.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => void }} Behavior */

const URL = 'http://host/ui/#/case/2007720723';
const SPEC = {sheets: [{name: 'Sheet1', cells: [{ref: 'A1', text: 'open case', hyperlink: URL}]}]};
// A foreign-authored file where the fragment ("#myhash") lives in the hyperlink element's location
// attribute, separate from the relationship Target (the bare "http://localhost/") — the read path
// must rejoin them, exercising the same fidelity from the READ direction rather than write→read.
const FIXTURE = 'hyperlink-url-fragment-preserved-on-round-trip/foreign-fragment.xlsx';

export default {
  id: 'hyperlink-url-fragment-preserved-on-round-trip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1084},
  cluster: 'xlsx-io',
  description:
    'A cell hyperlink whose URL contains a "#" fragment must round-trip in full, ' +
    'including the fragment, and preserve the display text.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a hyperlink URL with a fragment round-trips in full',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(sheets.Sheet1.cells.A1.hyperlink, URL);
      },
    },
    {
      name: 'the fragment tail is not silently dropped',
      baseline: 'pass',
      async expect(api, assert) {
        const {hyperlink} = (await api.roundtripWorkbook(SPEC)).sheets.Sheet1.cells.A1;
        assert.notStrictEqual(hyperlink, 'http://host/ui/', 'fragment was dropped');
        assert.ok(String(hyperlink).includes('/case/2007720723'), 'fragment segment missing');
      },
    },
    {
      name: 'the display text is preserved alongside the fragment URL',
      baseline: 'pass',
      async expect(api, assert) {
        const {text} = (await api.roundtripWorkbook(SPEC)).sheets.Sheet1.cells.A1;
        assert.strictEqual(text, 'open case');
      },
    },
    {
      name: 'reading a foreign file rejoins the fragment from the location attribute onto the base URL',
      baseline: 'pass',
      async expect(api, assert) {
        const links = await api.readFixtureHyperlinks(FIXTURE);
        assert.ok(links.A1, 'the hyperlink cell is read');
        assert.ok(
          String(links.A1.hyperlink).includes('#myhash'),
          `the fragment carried in the location attribute must be rejoined onto the base URL, not dropped; got ${JSON.stringify(links.A1)}`,
        );
      },
    },
  ],
};
