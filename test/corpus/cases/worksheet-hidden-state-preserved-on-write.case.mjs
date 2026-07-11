// Cluster: worksheet-decl
//
// Real-world scenario: a workbook has several worksheets, one marked hidden and another very hidden,
// while at least one stays visible (a valid workbook must keep one visible sheet). After writing and
// reopening, the marked worksheets must still be hidden: a consumer reading the file sees each sheet's
// visibility state, and the workbook's sheet-list entry declares that state rather than defaulting to
// visible. veryHidden — a state only settable through the file format, not the application UI — must
// not degrade to plain hidden or to visible.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'worksheet-hidden-state-preserved-on-write',
  provenance: {source: 'upstream-issue'},
  cluster: 'worksheet-decl',
  description:
    "A worksheet's visibility state (visible / hidden / veryHidden) survives a write: it reads back " +
    'unchanged and the workbook sheet-list entry carries the state attribute, with veryHidden not ' +
    'degrading to hidden or visible.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a hidden worksheet reads back hidden and a visible one stays visible',
      baseline: 'pass',
      async expect(api, assert) {
        const {readStates} = await api.worksheetStateReport();
        assert.strictEqual(readStates.Visible, 'visible', 'the visible sheet stays visible');
        assert.strictEqual(readStates.Hid, 'hidden', 'the hidden sheet reads back hidden');
      },
    },
    {
      name: 'a veryHidden worksheet reads back veryHidden, not degraded to hidden',
      baseline: 'pass',
      async expect(api, assert) {
        const {readStates} = await api.worksheetStateReport();
        assert.strictEqual(readStates.VeryHid, 'veryHidden', 'veryHidden must not degrade');
      },
    },
    {
      name: 'the workbook sheet-list entries declare each visibility state',
      baseline: 'pass',
      async expect(api, assert) {
        const {xmlStates} = await api.worksheetStateReport();
        assert.strictEqual(xmlStates.Hid, 'hidden', 'the hidden sheet entry carries state="hidden"');
        assert.strictEqual(xmlStates.VeryHid, 'veryHidden', 'the veryHidden sheet entry carries state="veryHidden"');
      },
    },
  ],
};
