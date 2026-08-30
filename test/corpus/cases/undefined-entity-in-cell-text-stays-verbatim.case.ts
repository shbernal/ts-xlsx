// Cluster: security
//
// Real-world scenario: an OOXML package declares no DTD, so a `&name;` the XML specification
// does not predefine has nothing to expand to. A reader that invents an expansion for it is
// the first half of an entity-expansion attack; a reader that answers it from whatever its
// host language happens to have lying around is worse, because the answer is not even text.
// A JavaScript object literal used as an entity table answers `&constructor;`, `&toString;`
// and a dozen more names with a function, and that function then flows into the cell value a
// caller reads and into the file the writer re-emits. The only correct reading of an entity
// no DTD defines is the characters the file actually spells.

import type {Assert, Case, CorpusApi} from '../case.ts';

// The names JavaScript's own prototype chain would answer, plus one that nothing answers, so a
// pass proves the miss path is reached rather than that the table happens to be empty.
const NAMES = ['constructor', 'toString', 'valueOf', 'hasOwnProperty', 'isPrototypeOf', 'lol'];

export default {
  id: 'undefined-entity-in-cell-text-stays-verbatim',
  cluster: 'security',
  description:
    'A cell whose text carries an entity reference no DTD defines reads back as the literal ' +
    'characters of that reference, and reads back as a string. This holds for the names a ' +
    'JavaScript object inherits from Object.prototype (`&constructor;`, `&toString;`) exactly ' +
    'as it holds for an invented one: the package has no DTD, so refusing to expand is what ' +
    'makes entity-expansion attacks structurally impossible rather than merely mitigated.',
  provenance: {source: 'hostile-input-review'},
  behavior: [
    {
      name: 'an entity no DTD defines reads back as the literal characters of the reference',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.unknownEntityReport(NAMES);
        for (const name of NAMES) {
          assert.strictEqual(report[name]?.value, `&${name};`, `&${name}; must not be expanded`);
        }
      },
    },
    {
      name: 'the reader never answers an undefined entity out of its own runtime',
      expect(api: CorpusApi, assert: Assert) {
        // The failure this guards is not a wrong character: it is the reader's host language
        // volunteering a value, which arrives as a function or as that function's source text.
        const report = api.unknownEntityReport(NAMES);
        for (const name of NAMES) {
          const {type, value} = report[name] ?? {};
          assert.strictEqual(type, 'string', `&${name}; must decode to a string`);
          assert.doesNotMatch(
            typeof value === 'string' ? value : '',
            /native code|\[object /,
            `&${name}; must not carry runtime internals into the cell`,
          );
        }
      },
    },
  ],
} satisfies Case;
