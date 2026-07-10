// Cluster: csv
//
// Real-world scenario: a user exports a CSV containing non-ASCII text (here Hebrew) and opens it in a
// spreadsheet application. The characters come out garbled because the file has no encoding marker:
// the bytes are valid UTF-8, but without a UTF-8 byte-order mark many spreadsheet apps (Excel most
// notably) fall back to a legacy code page and mis-decode the text. A CSV writer that already emits
// UTF-8 bytes should also emit the UTF-8 BOM so the encoding is detected, while the underlying text
// bytes must of course still decode back to the original string.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'csv-non-ascii-carries-utf8-bom',
  provenance: {source: 'upstream-issue'},
  cluster: 'csv',
  description:
    'A CSV containing non-ASCII text is written as UTF-8 with a byte-order mark so spreadsheet apps ' +
    'detect the encoding and render the characters correctly, and the bytes decode back to the ' +
    'original text.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the non-ASCII text is encoded as recoverable UTF-8 bytes',
      baseline: 'pass',
      async expect(api, assert) {
        const {bytesDecodeToText} = await api.csvNonAsciiEncodingReport();
        assert.strictEqual(bytesDecodeToText, true, 'the UTF-8 bytes decode back to the original text');
      },
    },
    {
      name: 'the CSV carries a UTF-8 BOM so the encoding is detected',
      baseline: 'fail',
      async expect(api, assert) {
        const {hasBom} = await api.csvNonAsciiEncodingReport();
        assert.strictEqual(hasBom, true, 'the CSV output begins with a UTF-8 byte-order mark');
      },
    },
  ],
};
