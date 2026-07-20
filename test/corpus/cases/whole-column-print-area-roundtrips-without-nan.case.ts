// Cluster: tables
//
// Real-world scenario: a worksheet's print area is set to span entire columns (or entire rows) —
// e.g. "print columns A through D, all rows". In the package this is a `_xlnm.Print_Area` defined
// name whose value is a column-only reference like 'S'!$A:$D, which is a perfectly valid OOXML
// range. On write the library emits exactly that. The bug is on the READ path: the reader decodes
// each endpoint as a full cell address and, finding no row number on a column-only reference,
// substitutes NaN — so reopening the file surfaces the print area back as the corrupt string
// "ANaN:DNaN". A caller inspecting or re-emitting that value gets a malformed reference that a
// spreadsheet app rejects. A bounded rectangular range (A1:D10) has both a column and a row at each
// endpoint, so it round-trips unmangled — which isolates the defect to column-/row-only references.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'whole-column-print-area-roundtrips-without-nan',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A whole-column (or whole-row) print area is a valid column-only OOXML reference (e.g. $A:$D); ' +
    'reading the written file back must recover that reference intact, not decode the missing row ' +
    'bound as NaN and surface a corrupt "ANaN:DNaN" address. A bounded rectangular print area, ' +
    'which carries a row at each endpoint, round-trips cleanly either way.',

  behavior: [
    {
      name: 'a whole-column print area writes a valid column-only defined name',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writtenDefinedName} = await api.printAreaRoundtrip('A:D');
        assert.ok(
          /\$A:\$D$/.test(String(writtenDefinedName)),
          `the written Print_Area name is the column-only reference $A:$D; got ${JSON.stringify(writtenDefinedName)}`,
        );
      },
    },
    {
      name: 'reading a whole-column print area back does not leak NaN into the recovered address',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reReadPrintArea} = await api.printAreaRoundtrip('A:D');
        assert.ok(
          !/NaN/.test(String(reReadPrintArea)),
          `a column-only print area must round-trip without NaN; got ${JSON.stringify(reReadPrintArea)}`,
        );
      },
    },
    {
      name: 'a bounded rectangular print area round-trips unmangled (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reReadPrintArea, reloadOk} = await api.printAreaRoundtrip('A1:D10');
        assert.strictEqual(reloadOk, true, 'the workbook reloads');
        assert.strictEqual(
          reReadPrintArea,
          'A1:D10',
          'a bounded range survives the round-trip exactly',
        );
      },
    },
  ],
} satisfies Case;
