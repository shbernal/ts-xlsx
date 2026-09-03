// Cluster: security
//
// Real-world scenario: a `.xlsx` this library did not write puts `&#1;` or `&#xD800;` in a sheet
// name, or a raw control byte in a defined name's formula. XML 1.0 has no representation for any of
// them -- `&#1;` is not an escape for U+0001, it is another way of spelling an ill-formed document --
// and the writer refuses all three, correctly. The reader's bound was wider than the writer's, so it
// decoded them into the model and the *next write* threw `AuthoringError: cannot write U+0001 at
// offset 2`: a hostile file reported as the caller's mistake, several layers and one taxonomy
// boundary from the input that caused it.
//
// The rule this locks, and it is the general one: the reader may only produce values the writer can
// serialise. A load-and-resave of a hostile file is not allowed to fail, and no character the format
// cannot carry reaches the model by any route.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'the-reader-never-decodes-what-the-writer-must-refuse',
  provenance: {source: 'hostile-input-probe'},
  cluster: 'security',
  description:
    'A character reference naming a code point XML 1.0 cannot carry, and a raw one, are refused at ' +
    'the reader boundary rather than decoded into the model: the read succeeds, the model carries ' +
    'nothing unrepresentable, and writing it back out does not throw.',

  behavior: [
    {
      name: 'a character the format cannot carry never aborts the read',
      expect(api: CorpusApi, assert: Assert) {
        for (const row of api.unrepresentableCharacterReport()) {
          assert.equal(row.threw, false, `${row.where} / ${row.spelling}: the read survives it`);
        }
      },
    },
    {
      name: 'nothing unrepresentable reaches the model',
      expect(api: CorpusApi, assert: Assert) {
        for (const row of api.unrepresentableCharacterReport()) {
          assert.equal(
            row.carriesUnrepresentable,
            false,
            `${row.where} / ${row.spelling}: decoded into the model`,
          );
        }
      },
    },
    {
      name: 'a load-and-resave of the hostile file does not throw at the caller',
      expect(api: CorpusApi, assert: Assert) {
        for (const row of api.unrepresentableCharacterReport()) {
          assert.equal(row.rewrote, true, `${row.where} / ${row.spelling}: the model re-writes`);
        }
      },
    },
    {
      name: 'a reference keeps the spelling the file used; a raw character has none to keep',
      expect(api: CorpusApi, assert: Assert) {
        const rows = api.unrepresentableCharacterReport();
        const sheetNameFor = (spelling: string) =>
          rows.find((row) => row.where === 'sheet name' && row.spelling === spelling)?.sheetName;
        assert.equal(sheetNameFor('&#1;'), 'A&#1;lpha', 'the reference is left verbatim');
        assert.equal(sheetNameFor('&#xD800;'), 'A&#xD800;lpha', 'and so is a lone surrogate one');
        assert.equal(
          sheetNameFor('a raw control character'),
          'Alpha',
          'a raw character has no verbatim form, so it is dropped',
        );
      },
    },
  ],
} satisfies Case;
