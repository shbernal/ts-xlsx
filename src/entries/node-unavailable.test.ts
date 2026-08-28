import assert from 'node:assert/strict';
import {test} from 'node:test';

import {StreamedRow, WorkbookStreamWriter, WorksheetStreamWriter} from './node-unavailable.ts';

// This module is what a browser bundle links in place of `entries/node.ts`, so its whole job is to
// fail legibly. A stub that threw something anonymous would be no better than the `undefined is not
// a function` it exists to replace: the name and the way out both have to be in the message.
test('each streaming class throws by name, and says what to use instead', () => {
  for (const [name, construct] of [
    ['StreamedRow', () => new StreamedRow()],
    ['WorksheetStreamWriter', () => new WorksheetStreamWriter()],
    ['WorkbookStreamWriter', () => new WorkbookStreamWriter()],
  ] as const) {
    assert.throws(construct, (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, new RegExp(`\\b${name}\\b`));
      assert.match(error.message, /not available in this environment/);
      assert.match(error.message, /writeXlsx/);
      return true;
    });
  }
});
