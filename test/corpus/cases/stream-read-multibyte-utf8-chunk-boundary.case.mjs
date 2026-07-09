// Cluster: streaming
//
// Real-world scenario: a workbook holds cells with multi-byte UTF-8 text (CJK characters, emoji) in
// the shared-strings table. The streaming reader consumes the shared-strings XML as a sequence of
// raw byte chunks fed to a SAX parser. If a chunk boundary falls in the middle of a multi-byte UTF-8
// sequence and each chunk is decoded independently, the split code point becomes the Unicode
// replacement character (U+FFFD) — so some CJK/emoji cells read back as garbage even though the file
// is well-formed and round-trips correctly through the non-streaming reader. A correct streaming
// decoder reassembles UTF-8 sequences across chunk boundaries. The payload here is deliberately large
// so the shared-strings part spans multiple underlying chunks.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// Large enough that the shared-strings XML is split across several stream chunks.
const CJK = '中文测试数据'.repeat(4000);
const EMOJI = '😀🎉🚀'.repeat(4000);

const SPEC = {
  sheets: [{name: 'S', cells: [{ref: 'A1', value: CJK}, {ref: 'A2', value: EMOJI}]}],
};

export default {
  id: 'stream-read-multibyte-utf8-chunk-boundary',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'The streaming reader must return multi-byte UTF-8 cell text (CJK, emoji) byte-exact even when ' +
    'the shared-strings XML is split across chunk boundaries — no U+FFFD replacement characters — ' +
    'and match what the non-streaming reader returns. Today a chunk split mid-character corrupts ' +
    'some sequences: a large CJK payload comes back with replacement characters (known-open), while ' +
    'a large emoji payload happens to survive — the reader decodes chunks independently instead of ' +
    'reassembling UTF-8 across boundaries.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'large CJK text streams back exactly, with no replacement characters',
      baseline: 'fail',
      async expect(api, assert) {
        const {streamed} = await api.streamReadSpec(SPEC, ['A1']);
        assert.ok(!String(streamed.A1).includes('�'), 'no U+FFFD from a chunk split mid-character');
        assert.strictEqual(streamed.A1, CJK, 'the streamed CJK value equals the source exactly');
      },
    },
    {
      name: 'large emoji text streams back exactly, with no replacement characters',
      baseline: 'pass',
      async expect(api, assert) {
        const {streamed} = await api.streamReadSpec(SPEC, ['A2']);
        assert.ok(!String(streamed.A2).includes('�'), 'no U+FFFD in the emoji cell');
        assert.strictEqual(streamed.A2, EMOJI, 'the streamed emoji value equals the source exactly');
      },
    },
    {
      name: 'the streaming reader output matches the non-streaming reader for the same file',
      baseline: 'fail',
      async expect(api, assert) {
        const {streamed, eager} = await api.streamReadSpec(SPEC, ['A1', 'A2']);
        assert.strictEqual(streamed.A1, eager.A1, 'CJK cell: streaming matches eager');
        assert.strictEqual(streamed.A2, eager.A2, 'emoji cell: streaming matches eager');
      },
    },
  ],
};
