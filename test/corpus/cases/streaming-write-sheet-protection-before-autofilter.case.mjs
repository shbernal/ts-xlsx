// Cluster: streaming
//
// Real-world scenario: a worksheet written through the streaming writer carries both an autofilter
// over its header range and sheet protection. OOXML's CT_Worksheet sequence requires the
// <sheetProtection> element to appear before <autoFilter>; emitting them reversed makes strict
// consumers (Excel) treat the file as corrupt. The buffered writer orders them correctly, so this
// isolates the streaming path's sibling-element ordering — the same class as the streaming
// conditionalFormatting/dataValidations-before-hyperlinks defects.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-write-sheet-protection-before-autofilter',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'A streaming-written worksheet with both an autofilter and sheet protection emits ' +
    '<sheetProtection> before <autoFilter> (the CT_Worksheet order), producing a valid package ' +
    'rather than a corrupt one with the elements reversed.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the streamed package reloads (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadOk} = await api.streamAutoFilterProtectionOrder();
        assert.strictEqual(
          reloadOk,
          true,
          'the streamed workbook reloads through the tolerant reader',
        );
      },
    },
    {
      name: 'sheetProtection is emitted before autoFilter',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetProtectionBeforeAutoFilter} = await api.streamAutoFilterProtectionOrder();
        assert.strictEqual(
          sheetProtectionBeforeAutoFilter,
          true,
          '<sheetProtection> must precede <autoFilter> per CT_Worksheet; the streaming writer emits them reversed',
        );
      },
    },
  ],
};
