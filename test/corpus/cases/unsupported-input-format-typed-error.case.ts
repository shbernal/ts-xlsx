import type {Assert, Case, CorpusApi} from '../case.ts';

// A spreadsheet reader is handed untrusted files, and not all of them are `.xlsx`. A legacy binary
// `.xls`, a binary `.xlsb`, a CSV pointed at the wrong reader, or an outright corrupt archive must each
// fail with a clear, catchable, format-tagged error — never a raw zip-library crash (which is opaque and
// can leak an absolute filesystem path from the layer below). The reader classifies the input and throws
// a typed error carrying a `format` field a caller branches on. See the spec
// `docs/knowledge/specs/unsupported-input-format-typed-error.md`.
export default {
  id: 'unsupported-input-format-typed-error',
  cluster: 'xlsx-io',
  description:
    'Non-.xlsx input (legacy .xls, binary .xlsb, non-ZIP text, corrupt ZIP) is rejected with a ' +
    'typed, format-tagged error rather than raw zip internals or a leaked filesystem path, while a ' +
    'genuine .xlsx still reads.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'a genuine .xlsx still reads through the format probe (no false rejection)',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('xlsx');
        assert.equal(result.threw, false);
      },
    },
    {
      name: 'a legacy .xls (OLE2/CFB) is rejected as an UnsupportedFormatError with format "xls"',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('xls');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'xls');
        assert.match(result.message, /\.xls/);
      },
    },
    {
      name: 'a binary .xlsb is rejected as an UnsupportedFormatError with format "xlsb"',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('xlsb');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'xlsb');
        assert.match(result.message, /\.xlsb/);
      },
    },
    {
      name: 'non-ZIP text is rejected as an UnsupportedFormatError with format "unknown"',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('garbage');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'unknown');
        assert.match(result.message, /not a valid \.xlsx package/);
      },
    },
    {
      name: 'a ZIP-headed but corrupt archive fails typed, leaking no zip internals or filesystem path',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('corrupt-zip');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'unknown');
        assert.equal(result.leaksZipInternals, false);
        assert.equal(result.leaksAbsolutePath, false);
      },
    },
    {
      name: 'the streaming reader enforces the same typed-error contract for a legacy .xls',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyStreamReadInput('xls');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'xls');
      },
    },
  ],
} satisfies Case;
