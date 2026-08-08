import type {Assert, Case, CorpusApi} from '../case.ts';

// A spreadsheet reader is handed untrusted files, and not all of them are `.xlsx`. A legacy binary
// `.xls`, a binary `.xlsb`, a CSV pointed at the wrong reader, or an outright corrupt archive must each
// fail with a clear, catchable, typed error — never a raw zip-library crash (which is opaque and can
// leak an absolute filesystem path from the layer below). Two branches, and the distinction is the
// contract: input of the wrong *kind* is an unsupported format (tagged with which one), while a
// container we recognise and cannot unpack is a malformed package. See the spec
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
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('xlsx');
        assert.equal(result.threw, false);
      },
    },
    {
      name: 'a legacy .xls (OLE2/CFB) is rejected as an UnsupportedFormatError with format "xls"',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('xls');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'xls');
        assert.match(result.message, /\.xls/);
      },
    },
    {
      name: 'a package whose office document is a binary workbook is parsed as one, not rejected',
      expect(api: CorpusApi, assert: Assert) {
        // The reader classifies this as an `.xlsb` and hands it to the BIFF12 codec, so the failure
        // it reports for a deliberately unparseable binary workbook is a *parse* error — the format
        // was recognised. (That a well-formed `.xlsb` reads into the same model its `.xlsx` twin does
        // is the subject of `xlsb-binary-workbook-reads-like-its-xlsx-twin`.)
        const result = api.classifyReadInput('xlsb');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'XlsbParseError');
        assert.equal(result.format, null);
      },
    },
    {
      name: 'the row streamer reports a binary .xlsb as a format it cannot take, naming one that can',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyStreamReadInput('xlsb');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'xlsb');
        assert.match(result.message, /\.xlsb/);
        assert.match(result.message, /readXlsx|readXlsb/);
      },
    },
    {
      name: 'non-ZIP text is rejected as an UnsupportedFormatError with format "unknown"',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('garbage');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'unknown');
        assert.match(result.message, /not a valid \.xlsx package/);
        // The sniff refused before the archive was opened, so the message must not blame a missing
        // part: it would name a check that never ran.
        assert.doesNotMatch(result.message, /workbook part/);
      },
    },
    {
      name: 'a corrupt archive is reported as a package that cannot be unpacked, not an unknown format',
      expect(api: CorpusApi, assert: Assert) {
        // A truncated package is the right *kind* of container; nothing inflated, so no part search
        // ever ran. Reporting it as an unrecognised format would name a check that did not happen and
        // point an investigation a layer past the one that refused.
        const result = api.classifyReadInput('corrupt-zip');
        assert.equal(result.threw, true);
        assert.equal(result.code, 'malformed-input');
        assert.notEqual(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, null);
      },
    },
    {
      name: 'a corrupt archive leaks neither zip internals nor a filesystem path',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyReadInput('corrupt-zip');
        assert.equal(result.threw, true);
        assert.equal(result.leaksZipInternals, false);
        assert.equal(result.leaksAbsolutePath, false);
      },
    },
    {
      name: 'the streaming reader classifies a corrupt archive the same way',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyStreamReadInput('corrupt-zip');
        assert.equal(result.threw, true);
        assert.equal(result.code, 'malformed-input');
        assert.equal(result.leaksZipInternals, false);
        assert.equal(result.leaksAbsolutePath, false);
      },
    },
    {
      name: 'the streaming reader enforces the same typed-error contract for a legacy .xls',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.classifyStreamReadInput('xls');
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'UnsupportedFormatError');
        assert.equal(result.format, 'xls');
      },
    },
  ],
} satisfies Case;
