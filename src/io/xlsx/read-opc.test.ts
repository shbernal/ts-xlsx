import assert from 'node:assert/strict';
import {test} from 'node:test';

import {resolveRelativePart, resolveWorkbookPart} from './read-opc.ts';

// Path resolution is a hostile-input parser path: a relationship Target comes verbatim from an
// untrusted package. These pin the OPC-legal shapes a well-formed writer never emits — absolute
// (package-root) targets and `.`/`..`/empty segments — so a real or malicious file that uses them
// still resolves to a bounded part path.

test('resolveRelativePart treats a leading slash as package-root-absolute', () => {
  assert.strictEqual(
    resolveRelativePart('xl/worksheets/sheet1.xml', '/xl/media/image1.png'),
    'xl/media/image1.png',
    'the base directory is ignored and the leading slash is stripped',
  );
});

test('resolveRelativePart collapses `.` and `..` segments against the base directory', () => {
  assert.strictEqual(
    resolveRelativePart('xl/worksheets/sheet1.xml', '../media/./image1.png'),
    'xl/media/image1.png',
    '`..` pops the parent and `.` is dropped',
  );
});

test('resolveRelativePart drops empty segments from a doubled slash', () => {
  assert.strictEqual(
    resolveRelativePart('xl/drawings/drawing1.xml', 'sub//child.xml'),
    'xl/drawings/sub/child.xml',
    'the empty segment between the slashes is skipped',
  );
});

test('resolveWorkbookPart roots a relative target under `xl/`', () => {
  assert.strictEqual(
    resolveWorkbookPart('worksheets/sheet1.xml'),
    'xl/worksheets/sheet1.xml',
    'a workbook-relative target is prefixed with the xl directory',
  );
});

test('resolveWorkbookPart treats a leading slash as package-root-absolute', () => {
  assert.strictEqual(
    resolveWorkbookPart('/xl/styles.xml'),
    'xl/styles.xml',
    'an absolute target is not prefixed, only de-slashed',
  );
});

test('resolveWorkbookPart strips a leading `./` before rooting under `xl/`', () => {
  assert.strictEqual(
    resolveWorkbookPart('./styles.xml'),
    'xl/styles.xml',
    'the current-directory prefix does not double the xl segment',
  );
});
