// Cluster: xlsx-io
//
// Real-world scenario: a user creates a workbook and adds worksheets without specifying a visibility
// state, then writes the file. The sheets must open as normal visible tabs. A regression once caused
// newly added sheets to be emitted as hidden unless the caller explicitly passed a visible state —
// the opposite of the natural expectation. Each sheet declaration in the workbook part carries a
// visibility marker (visible / hidden / veryHidden); with no state given, the written state must be
// visible. Explicit hidden/visible states must be preserved.

import type {Assert, Case, CorpusApi} from '../case.ts';
import type {Untyped} from '../untyped.ts';

const SPEC = {
  sheets: [
    {name: 'Default'}, // no state → must be visible
    {name: 'Hidden', state: 'hidden'},
    {name: 'Shown', state: 'visible'},
  ],
};

const stateOf = (entries: Untyped, name: Untyped) => {
  const e = entries.find((s: Untyped) => s.name === name);
  return e ? e.state : undefined;
};

export default {
  id: 'added-worksheet-defaults-to-visible',
  provenance: {source: 'upstream-issue'},
  cluster: 'xlsx-io',
  description:
    'A worksheet added with no visibility state is written as visible (never hidden/veryHidden), and ' +
    'explicit hidden or visible states are preserved — so default sheets open as normal tabs.',

  behavior: [
    {
      name: 'a sheet added with no state is visible (not hidden or veryHidden)',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetEntries} = await api.inspectPackage(SPEC);
        const state = stateOf(sheetEntries, 'Default');
        assert.ok(
          state === 'visible' || state === null,
          `a default-added sheet must be visible; got state ${JSON.stringify(state)}`,
        );
      },
    },
    {
      name: 'an explicitly hidden sheet is preserved as hidden',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetEntries} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          stateOf(sheetEntries, 'Hidden'),
          'hidden',
          'the hidden state survives to the workbook declaration',
        );
      },
    },
    {
      name: 'an explicitly visible sheet is preserved as visible',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetEntries} = await api.inspectPackage(SPEC);
        const state = stateOf(sheetEntries, 'Shown');
        assert.ok(
          state === 'visible' || state === null,
          `an explicitly visible sheet stays visible; got ${JSON.stringify(state)}`,
        );
      },
    },
    {
      name: 'no default-added sheet is emitted as veryHidden',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetEntries} = await api.inspectPackage(SPEC);
        assert.ok(
          sheetEntries.every((s) => s.state !== 'veryHidden'),
          'no sheet is silently marked veryHidden',
        );
      },
    },
  ],
} satisfies Case;
