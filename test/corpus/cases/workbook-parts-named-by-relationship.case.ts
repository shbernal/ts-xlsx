// Cluster: xlsx-io
//
// Real-world scenario: `xl/workbook.xml`, `xl/sharedStrings.xml` and `xl/styles.xml` are where Excel
// puts those parts, not where OPC says they live. A package names its office document in
// `_rels/.rels` and its pool and stylesheet in the workbook's own `.rels`, and a conforming producer
// is free to point those relationships anywhere: exporters that lay their packages out by feature,
// tools that rewrite a package part by part, and hand-assembled OPC all do. Excel reads such a file
// without comment, because it follows the graph.
//
// The two halves fail differently, which is why they are two fixtures. A workbook part the reader
// cannot find is loud: nothing opens. A pool or stylesheet it cannot find is silent and worse: every
// `t="s"` cell reads as the empty string, and an empty style table means no cell has a number format,
// so a date serial comes back as the bare number 45000 rather than a date. That is a change of cell
// *type*, with no error anywhere.

import type {Assert, Case, CorpusApi} from '../case.ts';

const RENAMED_ALL = 'workbook-parts-named-by-relationship/renamed-parts.xlsx';
const RENAMED_PARTS_ONLY = 'workbook-parts-named-by-relationship/renamed-pool-and-styles.xlsx';

export default {
  id: 'workbook-parts-named-by-relationship',
  provenance: {source: 'foreign-generator-probe'},
  cluster: 'xlsx-io',
  description:
    'The office document, the shared-string pool and the stylesheet are resolved through the ' +
    'package relationship graph, not by their conventional paths: a package that names them ' +
    'elsewhere reads its sheets, its pooled strings and its date cells exactly as one that does not.',

  behavior: [
    {
      name: 'a workbook part named somewhere other than xl/workbook.xml still opens',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.readFixtureReport(RENAMED_ALL);
        assert.ok(report.ok, `the package must open; got ${String(report.error)}`);
        assert.deepStrictEqual(report.sheetNames, ['S']);
      },
    },
    {
      name: 'a pool named somewhere other than xl/sharedStrings.xml still resolves its strings',
      async expect(api: CorpusApi, assert: Assert) {
        for (const fixture of [RENAMED_ALL, RENAMED_PARTS_ONLY]) {
          const cells = await api.readFixtureCells(fixture, ['A1']);
          assert.strictEqual(
            cells.A1.value,
            'pooled text',
            `a t="s" cell must resolve through the pool the workbook points at, in ${fixture}`,
          );
        }
      },
    },
    {
      name: 'a stylesheet named somewhere other than xl/styles.xml still makes a serial a date',
      async expect(api: CorpusApi, assert: Assert) {
        for (const fixture of [RENAMED_ALL, RENAMED_PARTS_ONLY]) {
          const cells = await api.readFixtureCells(fixture, ['A2']);
          assert.strictEqual(
            cells.A2.type,
            'date',
            `an unresolved stylesheet changes the cell's type, not just its look, in ${fixture}`,
          );
        }
      },
    },
    {
      name: 'the streaming reader resolves the same parts the buffered reader does',
      async expect(api: CorpusApi, assert: Assert) {
        // Two readers, one package layout: a streamed read that resolved a part differently from a
        // buffered one would decode the same cell to a different value, which is the divergence the
        // shared accumulator exists to prevent.
        const streamed = await api.streamReadFixture(RENAMED_ALL, ['A1', 'A2']);
        assert.strictEqual(streamed.A1.value, 'pooled text');
        assert.strictEqual(streamed.A2.type, 'date');
      },
    },
  ],
} satisfies Case;
