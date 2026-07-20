// Cluster: streaming
//
// Real-world scenario: a worksheet written through the streaming writer has both a hyperlink cell and
// a cell with a data validation (a list dropdown). OOXML's CT_Worksheet content model fixes the
// order of trailing child elements — <dataValidations> must precede <hyperlinks>. The streaming
// writer emits the hyperlinks block before the data-validations block, so Microsoft Excel reports
// the file as corrupt and offers to recover it, while tolerant apps open it fine. The non-streaming
// writer orders them correctly. This is the same streaming child-order defect as the
// conditional-formatting/hyperlinks case, on the data-validation/hyperlink pair.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'streamed-worksheet-datavalidations-before-hyperlinks',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'A streamed worksheet carrying both a data validation and a hyperlink emits the <dataValidations> ' +
    'block before <hyperlinks>, per the CT_Worksheet sequence, so strict consumers do not treat the ' +
    'file as corrupt.',

  behavior: [
    {
      name: 'both blocks are present in the streamed worksheet XML',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {posDataValidations, posHyperlinks} = await api.streamWriteDvHyperlinkOrder();
        assert.ok(posDataValidations >= 0, 'a dataValidations block is emitted');
        assert.ok(posHyperlinks >= 0, 'a hyperlinks block is emitted');
      },
    },
    {
      name: 'dataValidations is emitted before hyperlinks (CT_Worksheet sequence order)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {dataValidationsBeforeHyperlinks} = await api.streamWriteDvHyperlinkOrder();
        assert.strictEqual(
          dataValidationsBeforeHyperlinks,
          true,
          'dataValidations must precede hyperlinks in the worksheet XML, or Excel repairs the file',
        );
      },
    },
    {
      name: 'the streamed package still reloads with the tolerant reader',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reloadOk} = await api.streamWriteDvHyperlinkOrder();
        assert.strictEqual(
          reloadOk,
          true,
          'the tolerant reader reads it back despite the wrong order',
        );
      },
    },
  ],
} satisfies Case;
