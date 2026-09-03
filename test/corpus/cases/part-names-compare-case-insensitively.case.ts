// Cluster: xlsx-io
//
// Real-world scenario: OPC compares part names case-insensitively (ASCII), so `/XL/Workbook.XML` and
// `/xl/workbook.xml` name one part. Producers take the format at its word: packages come out of
// generators, archivers and case-insensitive filesystems with their parts cased differently from the
// relationships and content-type overrides that reference them, and Excel opens them.
//
// A reader that compares part paths as raw strings finds nothing at all in such a package: the fold
// is not a nicety at the edges but the difference between reading the file and reporting it as
// unrecognised. The fixture cases every part reference differently from the entry it names.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'part-names-compare-case-insensitively/mixed-case.xlsx';

export default {
  id: 'part-names-compare-case-insensitively',
  provenance: {source: 'foreign-generator-probe'},
  cluster: 'xlsx-io',
  description:
    'Package part lookups fold ASCII case, as OPC requires: a package whose parts are cased ' +
    'differently from the relationships and overrides naming them reads its workbook, its ' +
    'worksheets, its pool and its stylesheet, rather than reporting no workbook part.',

  behavior: [
    {
      name: 'a workbook part cased differently from the relationship naming it still opens',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.readFixtureReport(FIXTURE);
        assert.ok(report.ok, `the package must open; got ${String(report.error)}`);
        assert.deepStrictEqual(report.sheetNames, ['S']);
      },
    },
    {
      name: 'the worksheet, pool and stylesheet resolve through the same fold',
      async expect(api: CorpusApi, assert: Assert) {
        // One assertion per part would be three ways of saying the same thing; what matters is that
        // the fold reaches every hop of the graph, not just the first. A cell value proves the sheet
        // and the pool; the cell's type proves the stylesheet.
        const cells = await api.readFixtureCells(FIXTURE, ['A1', 'A2', 'A3']);
        assert.strictEqual(cells.A1.value, 'pooled text');
        assert.strictEqual(cells.A2.type, 'date');
        assert.strictEqual(cells.A3.value, 42);
      },
    },
    {
      name: 'the streaming reader folds case the same way the buffered reader does',
      async expect(api: CorpusApi, assert: Assert) {
        const streamed = await api.streamReadFixture(FIXTURE, ['A1', 'A2']);
        assert.strictEqual(streamed.A1.value, 'pooled text');
        assert.strictEqual(streamed.A2.type, 'date');
      },
    },
  ],
} satisfies Case;
