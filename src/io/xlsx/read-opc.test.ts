import assert from 'node:assert/strict';
import {test} from 'node:test';
import {strToU8} from 'fflate';

import {
  capturePartClosure,
  packageAccessors,
  resolveRelativePart,
  resolveWorkbookPart,
} from './read-opc.ts';

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

// An externalLink part points at its source workbook through a `TargetMode="External"` relationship.
// The closure must keep that wiring verbatim — dropping it (as it once did) orphans the link and
// dangles every `[n]` external reference a formula resolves through — while never trying to walk into
// the out-of-package target.
test('capturePartClosure retains an external relationship verbatim without walking it', () => {
  const rels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" ' +
    'Target="C:\\Sources\\Linked.xlsm" TargetMode="External"/></Relationships>';
  const files = {
    'xl/externalLinks/externalLink1.xml': strToU8('<externalLink/>'),
    'xl/externalLinks/_rels/externalLink1.xml.rels': strToU8(rels),
  };
  const {partText, partBytes} = packageAccessors(files);
  const closure = capturePartClosure(
    'xl/externalLinks/externalLink1.xml',
    partText,
    partBytes,
    () => 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
  );

  assert.ok(closure !== undefined, 'the entry part is present');
  assert.strictEqual(closure.length, 1, 'the external target is not visited as a package part');
  const [entry] = closure;
  assert.deepStrictEqual(
    entry?.rels,
    [
      {
        id: 'rId1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
        targetPath: 'C:\\Sources\\Linked.xlsm',
        external: true,
      },
    ],
    'the external relationship is kept with its raw target and the external flag',
  );
});
