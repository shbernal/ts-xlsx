// Cluster: styles
//
// Real-world scenario: an Excel-authored workbook stores conditional formatting in the x14
// extension-list block of the worksheet (<extLst> → <x14:conditionalFormattings>) rather than the
// classic <conditionalFormatting> element. An expression-type rule there carries its formula inside
// an <xm:f> element. Loading such a workbook succeeds, but writing it back out crashes with an
// undefined-property access during conditional-formatting rendering, because the extension rule's
// formula is not wired into the write path. Reading and re-writing must not lose the rule or crash.
//
// The fixture is a workbook with an x14 extension-list expression rule (formula "A1>2" over A1:A5)
// injected into the worksheet XML.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'extlst-conditional-formatting/sample.xlsx';

export default {
  id: 'extlst-conditional-formatting-roundtrip-does-not-crash',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A workbook whose conditional formatting lives only in the x14 extension list loads and writes ' +
    'back without the writer crashing on the extension rule’s formula.',

  behavior: [
    {
      name: 'a workbook with x14 extension-list conditional formatting loads without throwing',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {loadOk, loadError} = await api.roundtripFixtureWriteReport(FIXTURE);
        assert.strictEqual(loadOk, true, `the load must succeed; got ${JSON.stringify(loadError)}`);
      },
    },
    {
      name: 'writing the workbook back out does not crash on the extension rule',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writeOk, writeError} = await api.roundtripFixtureWriteReport(FIXTURE);
        assert.strictEqual(
          writeOk,
          true,
          `the write-back must not throw; got ${JSON.stringify(writeError)}`,
        );
      },
    },
  ],
} satisfies Case;
