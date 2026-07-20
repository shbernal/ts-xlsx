// Cluster: xlsx-io
//
// Real-world scenario: code builds a workbook conditionally and, on an empty data
// set, ends up adding no worksheets before writing. A workbook with zero worksheets
// cannot be represented as a valid .xlsx — Excel shows the "unreadable content"
// repair prompt. A correct writer must resolve this unambiguously: either refuse
// (throw a clear error) or emit a package that actually contains a worksheet part.
// Silently producing a package that declares no worksheet is the wrong outcome.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'empty-workbook-no-worksheets',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 739},
  cluster: 'xlsx-io',
  description:
    'Writing a workbook with no worksheets must not silently produce a package that ' +
    'declares zero worksheet parts; it must either throw or emit a valid worksheet.',

  behavior: [
    {
      name: 'writing a workbook with no worksheets does not silently produce a corrupt package',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const write = await api.tryWriteWorkbook({sheets: []});
        if (!write.ok) return; // throwing is an acceptable resolution
        const {consistency} = await api.inspectPackage({sheets: []});
        assert.ok(
          consistency.worksheetPartCount >= 1,
          'write succeeded but the package declares zero worksheet parts (Excel-corrupt)',
        );
      },
    },
    {
      name: 'a one-worksheet workbook declares that worksheet consistently across the package',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {consistency} = await api.inspectPackage({
          sheets: [{name: 'Sheet1', cells: [{ref: 'A1', value: 'x'}]}],
        });
        assert.strictEqual(consistency.worksheetPartCount, 1);
        assert.strictEqual(consistency.sheetEntryCount, 1);
        assert.ok(
          consistency.declaredConsistent,
          'worksheet part not declared consistently in content-types + rels',
        );
      },
    },
  ],
} satisfies Case;
