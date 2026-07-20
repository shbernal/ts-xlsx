// Cluster: csv
//
// Real-world scenario: the CSV reader applies a per-value transformation to every parsed cell. By
// default it coerces — a numeric-looking cell like "007" becomes the number 7 (losing the leading
// zero), "32.5" becomes 32.5. A caller who needs the raw text (identifiers, leading zeros) — or who
// wants to skip the expensive default coercion on a large file — supplies a custom map function; an
// identity map preserves each raw string verbatim. The map option must govern how each parsed value
// is transformed before it becomes a cell.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'csv-read-map-option-controls-value-coercion',
  provenance: {source: 'upstream-issue'},
  cluster: 'csv',
  description:
    'The CSV read map option controls per-value transformation: an identity map preserves raw ' +
    'string values verbatim (leading zeros kept, no coercion) while the default map coerces ' +
    'numeric-looking cells to numbers.',

  behavior: [
    {
      name: 'the default map coerces a numeric-looking cell to a number',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {default: def} = await api.csvReadMapReport();
        assert.strictEqual(def.aType, 'number', 'the default coerces "007" to a number');
        assert.strictEqual(def.a, 7, 'the leading zero is lost under default coercion');
      },
    },
    {
      name: 'an identity map preserves the raw string value with its leading zero',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {identity} = await api.csvReadMapReport();
        assert.strictEqual(identity.aType, 'string', 'the identity map keeps the value a string');
        assert.strictEqual(
          identity.a,
          '007',
          'the raw text, including the leading zero, is preserved',
        );
        assert.strictEqual(identity.b, '32.5', 'the decimal string is preserved verbatim too');
      },
    },
  ],
} satisfies Case;
