// Cluster: styles
//
// Real-world scenario: a user attaches a multi-line note (cell comment) to a cell. When the file is
// opened, the note popup renders at a fixed default size too small to show the text, and the content
// is clipped. The cause is in the comment's VML drawing: its `<v:textbox>` style declares only
// `mso-direction-alt:auto` and omits `mso-fit-shape-to-text:t`, the directive that tells the host
// application to grow the note shape to fit its text. Users worked around it by post-processing the
// output zip to inject the directive into every comment textbox. Emitting `mso-fit-shape-to-text:t`
// so the note box sizes to its content is the correct behavior.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [{name: 'S', cells: [{ref: 'B2', value: 'x', note: 'line one\nline two\nline three'}]}],
};

export default {
  id: 'comment-note-box-fits-its-text',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A cell note is emitted with a VML textbox that auto-fits its text (mso-fit-shape-to-text:t), so ' +
    'a multi-line note is not clipped to a fixed default box; the comment VML part is present and ' +
    'well-formed.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the note is written with a VML drawing part',
      baseline: 'pass',
      async expect(api, assert) {
        const {packageParts, vml} = await api.inspectPackage(SPEC);
        assert.ok(packageParts.hasVmlDrawingPart, 'a comment writes its legacy VML drawing');
        assert.ok(vml.textboxStyles.length >= 1, 'the VML drawing declares a comment textbox');
      },
    },
    {
      name: 'the comment textbox is styled to auto-fit its text so a multi-line note is not clipped',
      baseline: 'pass',
      async expect(api, assert) {
        const {vml} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          vml.allTextboxesFitToText,
          true,
          `every comment textbox must carry mso-fit-shape-to-text:t; got styles ${JSON.stringify(vml.textboxStyles)}`,
        );
      },
    },
  ],
};
