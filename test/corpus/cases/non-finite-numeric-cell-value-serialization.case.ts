// Cluster: types
//
// Real-world scenario: a caller computes a cell value and, through a bug (e.g. parsing a missing
// field), assigns a non-finite JavaScript number — NaN, Infinity, or -Infinity — as a numeric cell
// value. On write, a literal `NaN`/`Infinity` token is not valid content for a numeric `<v>` element,
// so Excel reports "unreadable content" and offers to recover the file. The library should never
// serialize a bare non-finite token into a numeric cell — it should coerce to something a strict
// OOXML consumer accepts (an empty/blank cell, a text cell, or an error value), keeping the package
// well-formed.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'non-finite-numeric-cell-value-serialization',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    'Writing a cell whose value is NaN, Infinity, or -Infinity does not emit a bare non-finite token ' +
    'into a numeric <v> element (which Excel treats as unreadable content); the produced package is ' +
    'well-formed and re-readable.',

  behavior: [
    {
      name: 'a NaN cell value does not serialize a bare "NaN" token',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasNonFiniteToken, token} = await api.nonFiniteCellReport('NaN');
        assert.strictEqual(
          hasNonFiniteToken,
          false,
          `a NaN value must not emit a bare token; got <v>${token}</v>`,
        );
      },
    },
    {
      name: 'an Infinity cell value does not serialize a bare "Infinity" token',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasNonFiniteToken, token} = await api.nonFiniteCellReport('Infinity');
        assert.strictEqual(
          hasNonFiniteToken,
          false,
          `an Infinity value must not emit a bare token; got <v>${token}</v>`,
        );
      },
    },
    {
      name: 'a -Infinity cell value does not serialize a bare "-Infinity" token',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasNonFiniteToken, token} = await api.nonFiniteCellReport('-Infinity');
        assert.strictEqual(
          hasNonFiniteToken,
          false,
          `a -Infinity value must not emit a bare token; got <v>${token}</v>`,
        );
      },
    },
  ],
} satisfies Case;
