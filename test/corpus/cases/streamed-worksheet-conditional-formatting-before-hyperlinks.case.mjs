// Cluster: streaming
//
// Real-world scenario: OOXML's worksheet part (CT_Worksheet) fixes the order of its child elements —
// among them, <conditionalFormatting> must precede <hyperlinks>. When a sheet is written through the
// streaming writer with both a conditional-formatting rule (e.g. alternating-row shading via
// MOD(ROW(),2)=0) and a hyperlink cell, the streaming writer emits these two blocks in the wrong
// order (hyperlinks before conditionalFormatting), violating the schema sequence — so Excel reports
// the file as corrupt and offers to repair it. The buffered writer orders them correctly; this is a
// streaming-path ordering defect.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streamed-worksheet-conditional-formatting-before-hyperlinks',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'A streamed worksheet carrying both a conditional-formatting rule and a hyperlink emits the ' +
    '<conditionalFormatting> block before <hyperlinks>, per the CT_Worksheet schema sequence, so the ' +
    'file is not treated as corrupt.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'both blocks are present in the streamed worksheet XML',
      baseline: 'pass',
      async expect(api, assert) {
        const {posConditionalFormatting, posHyperlinks} = await api.streamWriteCfHyperlinkOrder();
        assert.ok(posConditionalFormatting >= 0, 'the conditionalFormatting block is emitted');
        assert.ok(posHyperlinks >= 0, 'the hyperlinks block is emitted');
      },
    },
    {
      name: 'conditionalFormatting is emitted before hyperlinks (CT_Worksheet sequence order)',
      baseline: 'pass',
      async expect(api, assert) {
        const {conditionalFormattingBeforeHyperlinks} = await api.streamWriteCfHyperlinkOrder();
        assert.strictEqual(
          conditionalFormattingBeforeHyperlinks,
          true,
          'conditionalFormatting must precede hyperlinks in the worksheet XML, or Excel repairs the file'
        );
      },
    },
    {
      name: 'the streamed package still reloads with this library’s tolerant reader',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadOk} = await api.streamWriteCfHyperlinkOrder();
        assert.strictEqual(reloadOk, true, 'the tolerant reader reads it back even though the order is wrong');
      },
    },
  ],
};
