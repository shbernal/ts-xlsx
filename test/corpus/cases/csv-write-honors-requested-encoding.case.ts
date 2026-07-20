// Cluster: csv
//
// Real-world scenario: a caller exports a worksheet to CSV and passes an explicit output encoding,
// expecting the produced bytes to be in that encoding. The multibyte content itself (emoji, CJK,
// accented Latin) is not the problem — it survives a UTF-8 round-trip verbatim. The problem is that
// the CSV writer silently ignores the requested encoding and always emits UTF-8, so a caller who
// asked for a different encoding (because their downstream consumer expects it) gets bytes that
// decode to garbage ("mojibake") under the encoding they think they have. A requested output
// encoding must actually take effect, not be dropped on the floor.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'csv-write-honors-requested-encoding',
  provenance: {source: 'upstream-issue'},
  cluster: 'csv',
  description:
    'Multibyte CSV content (emoji, CJK) survives a write/read round-trip verbatim under the default ' +
    'UTF-8 path; and when the caller requests a specific output encoding, the produced bytes are in ' +
    'that encoding rather than silently emitted as UTF-8.',

  behavior: [
    {
      name: 'emoji and CJK survive a default UTF-8 CSV round-trip verbatim',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {emojiRoundtrips} = await api.csvWriteEncodingReport();
        assert.strictEqual(
          emojiRoundtrips,
          true,
          'astral (emoji) and CJK characters must round-trip byte-for-byte under UTF-8, not be substituted or truncated',
        );
      },
    },
    {
      name: 'a requested non-UTF-8 output encoding is actually applied',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {decodesAsRequested, decodesAsUtf8, requestedEncoding} =
          await api.csvWriteEncodingReport({encoding: 'utf16le'});
        assert.strictEqual(
          decodesAsRequested,
          true,
          `requesting ${requestedEncoding} must produce bytes that decode under that encoding; ` +
            `instead the encoding was ignored and the bytes are UTF-8 (decodesAsUtf8=${decodesAsUtf8})`,
        );
      },
    },
  ],
} satisfies Case;
