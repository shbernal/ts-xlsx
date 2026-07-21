import assert from 'node:assert/strict';
import {test} from 'node:test';

import {extensionOf} from './part-paths.ts';

// A part path comes verbatim from an untrusted package (both a rels Target and a raw media path),
// so this must only recognise a `.` after the last `/` — a dot in a directory segment is not a file
// extension, even when the filename itself has none.
test('extensionOf only recognises a `.` after the last `/`, not one in an earlier segment', () => {
  assert.strictEqual(extensionOf('xl/media/image1.png'), 'png');
  assert.strictEqual(
    extensionOf('xl/media.v2/image1'),
    '',
    'a dot in a directory segment is not a file extension, even if the filename has none',
  );
});
