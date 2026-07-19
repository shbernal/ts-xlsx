// Cluster: streaming
//
// Real-world scenario: a workbook holds a large number of worksheets (well over one hundred). A
// caller opens it with the streaming workbook reader and iterates, counting each worksheet as it
// arrives. Every worksheet must be emitted exactly once. In the reported failure the reader halts or
// drops entries partway through once the sheet count is large: a worksheet part is reached and parsed
// before the workbook model (from xl/workbook.xml) has been built, so the iteration ends early — or
// throws — and the caller silently sees fewer worksheets than the file contains. A small workbook
// works, masking the defect. The count of emitted worksheets must equal the number written,
// regardless of scale. (The desired policy is captured in the streaming-read-emits-all-worksheets
// spec note; this case locks it against a many-sheet file.)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-read-emits-all-worksheets-at-scale',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'The streaming reader emits every worksheet of a many-sheet workbook exactly once — a workbook ' +
    'with far more than 100 sheets streams its full sheet count, without truncating the tail or ' +
    'throwing when a worksheet part is reached before the workbook model is built.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a small workbook streams every worksheet (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {written, emitted, error} = await api.streamReadManySheets(3);
        assert.strictEqual(error, null, `small streaming read must not error; got ${error}`);
        assert.strictEqual(emitted, written, 'every worksheet of a small workbook is emitted');
      },
    },
    {
      name: 'a many-sheet workbook streams without the reader throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {error} = await api.streamReadManySheets(180);
        assert.strictEqual(error, null, `streaming a many-sheet workbook must not throw; got ${error}`);
      },
    },
    {
      name: 'the streamed worksheet count equals the number written, at scale',
      baseline: 'pass',
      async expect(api, assert) {
        const {written, emitted} = await api.streamReadManySheets(180);
        assert.strictEqual(emitted, written, `every worksheet must be emitted; wrote ${written}, streamed ${emitted}`);
      },
    },
  ],
};
