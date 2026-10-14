import assert from 'node:assert/strict';
import {test} from 'node:test';

import {structuredColumnReference} from './structured-reference.ts';

test("every other character on Microsoft's special list takes the double-bracket form unescaped", () => {
  const unescaped = ['\t', '\n', '\r', ',', ':', '.', '"', '{', '}', '$', '^', '&', '*', '+', '='];
  unescaped.push('-', '>', '<', '/', '\\', '!', '(', ')', '%', '?', '`', ';', '~', '_');
  for (const ch of unescaped) {
    assert.equal(structuredColumnReference('T', `a${ch}b`), `T[[a${ch}b]]`, JSON.stringify(ch));
  }
});
