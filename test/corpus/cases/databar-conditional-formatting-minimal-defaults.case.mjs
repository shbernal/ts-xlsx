// Cluster: styles
//
// Real-world scenario: a user adds a data-bar conditional-formatting rule with only the minimal
// options — a rule of type "dataBar" with a priority, no explicit conditional-formatting-value-
// objects (cfvo), and no bar color. Excel's own UI fills in a min/max cfvo range and a standard bar
// color when you apply a default data bar; the library should do the same and produce a valid file.
// The bug: with cfvo omitted the writer indexes into an absent cfvo collection and throws
// ("Cannot read properties of undefined (reading 'forEach')"), so the minimal, most natural call
// crashes. Supplying explicit cfvo and color is the working control.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'databar-conditional-formatting-minimal-defaults',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A dataBar conditional-formatting rule with only type + priority (no cfvo, no color) writes ' +
    'without throwing and gains default min/max cfvo anchors and a default bar color, the same as ' +
    'Excel’s default data bar — rather than crashing on the absent cfvo collection.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a dataBar rule with explicit cfvo and color writes (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {writeOk} = await api.authorConditionalFormatting({
          ref: 'A1:A3',
          rules: [
            {
              type: 'dataBar',
              priority: 1,
              cfvo: [{type: 'min'}, {type: 'max'}],
              color: {argb: 'FF638EC6'},
            },
          ],
        });
        assert.strictEqual(writeOk, true, 'a fully-specified dataBar writes');
      },
    },
    {
      name: 'a minimal dataBar rule (no cfvo, no color) writes without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {writeOk, writeError} = await api.authorConditionalFormatting({
          ref: 'A1:A3',
          rules: [{type: 'dataBar', priority: 1}],
        });
        assert.strictEqual(
          writeOk,
          true,
          `a minimal dataBar must default its cfvo/color, not crash; got ${JSON.stringify(writeError)}`,
        );
      },
    },
    {
      name: 'the minimal dataBar gains two default cfvo anchors',
      baseline: 'pass',
      async expect(api, assert) {
        const {xml} = await api.authorConditionalFormatting({
          ref: 'A1:A3',
          rules: [{type: 'dataBar', priority: 1}],
        });
        assert.ok(xml?.hasDataBar, 'a dataBar element is emitted');
        assert.strictEqual(xml.cfvoCount, 2, 'a default data bar carries a min and a max cfvo');
      },
    },
    {
      name: 'the minimal dataBar gains a default bar color',
      baseline: 'pass',
      async expect(api, assert) {
        const {xml} = await api.authorConditionalFormatting({
          ref: 'A1:A3',
          rules: [{type: 'dataBar', priority: 1}],
        });
        assert.ok(xml?.hasColor, 'a default data bar carries a bar color when none was supplied');
      },
    },
  ],
};
