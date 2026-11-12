# Errors

A spreadsheet library mostly reads files someone else wrote, so it has to fail clearly.
Every error this library raises on purpose extends `XlsxError` and carries a `code` naming
the kind of failure.

```ts
import {readXlsx, XlsxError} from '@shbernal/ts-xlsx';

try {
  readXlsx(new Uint8Array([1, 2, 3, 4]));
  console.log('that should not have worked');
} catch (error) {
  console.log(error instanceof XlsxError); // true
  console.log(error instanceof XlsxError ? error.code : 'not ours'); // 'unsupported-format'
}
```

`error instanceof XlsxError` tells you whether the library raised it. Anything else is a bug
in your code or your runtime, and the library does not wrap it.

## The four codes, and what you would do about each

There are four codes because there are four different things a caller does next.

| `code`                 | What happened                                      | What you do                                    |
| ---------------------- | -------------------------------------------------- | ---------------------------------------------- |
| `'unsupported-format'` | Not a container this library reads at all.         | Reject the file, or hand it to something else. |
| `'malformed-input'`    | A part we do read is corrupt or off-specification. | Reject the file. It is broken, or hostile.     |
| `'authoring'`          | Your code described a document that cannot exist.  | Fix the calling code.                          |
| `'internal'`           | An invariant the library maintains did not hold.   | Report it. The bug is ours.                    |

Every subclass fixes `code` to a literal, so the hierarchy is a discriminated union and
narrowing on `error.code` narrows the type.

```ts
import {readXlsx, XlsxError} from '@shbernal/ts-xlsx';

function classify(bytes: Uint8Array): string {
  try {
    readXlsx(bytes);
    return 'read';
  } catch (error) {
    if (!(error instanceof XlsxError)) throw error;
    switch (error.code) {
      case 'unsupported-format':
        return 'not a spreadsheet we read';
      case 'malformed-input':
        return 'broken or hostile';
      case 'authoring':
        return 'our bug';
      case 'internal':
        return 'their bug, please report it';
    }
  }
}

console.log(classify(new Uint8Array([0, 0, 0, 0]))); // 'not a spreadsheet we read'
```

## Telling apart the two ways a file can be wrong

`UnsupportedFormatError` says the input is a different _kind_ of thing. `PackageReadError`
says it is the right kind and cannot be unpacked. Keeping them apart answers
"should I try another reader, or reject this?". The first carries a `format` field, so you
never match on a message.

```ts
import {readXlsx, UnsupportedFormatError} from '@shbernal/ts-xlsx';

// The magic bytes of a legacy OLE2 .xls.
const legacy = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);

try {
  readXlsx(legacy);
  console.log('that should not have worked');
} catch (error) {
  console.log(error instanceof UnsupportedFormatError); // true
  console.log(error instanceof UnsupportedFormatError ? error.format : ''); // 'xls'
}
```

`format` is `'xls'`, `'xlsb'` or `'unknown'`. A message never carries a filesystem path or
the zip library's internals, so nothing from a lower layer leaks into your logs or into an
error you show a user.

## Errors that are your fault

`AuthoringError` is raised when the _document you described_ cannot exist: a workbook with no
worksheets, a table whose columns do not span its range, a merge overlapping another.

```ts
import {AuthoringError, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

try {
  writeXlsx(new Workbook()); // a package with no worksheets is corrupt
  console.log('that should not have worked');
} catch (error) {
  console.log(error instanceof AuthoringError); // true
  console.log(error instanceof AuthoringError ? error.code : ''); // 'authoring'
}
```

A single argument that is out of range or the wrong type stays a native `RangeError` or
`TypeError`. JavaScript already has those types for exactly that case. `getColumn(0)` is a
`RangeError`; a table that names a column twice is an `AuthoringError`.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('S');
try {
  sheet.getColumn(0);
  console.log('that should not have worked');
} catch (error) {
  console.log(error instanceof RangeError); // true
}
```

## `InternalError` is not a failure mode to handle

It means an invariant this library maintains did not hold. No caller can provoke one, so
there is nothing to catch it for. It has its own type so that "this should be unreachable" is explicit, and it is the only
class that appends a "please report this" notice to its message. The others leave the message
as given. A corrupt input file is routine for a library that reads untrusted files, and a
report-this banner on each one would teach you to ignore the one that matters.

## Where each class is documented

The reference pages are generated from the types:

- [`errors`](../api/errors.md) for `XlsxError`, `AuthoringError` and `InternalError`;
- [`opc-errors`](../api/opc-errors.md) for `UnsupportedFormatError` and `PackageReadError`;
- [`xlsx-errors`](../api/xlsx-errors.md), [`xlsb-errors`](../api/xlsb-errors.md) and
  [`xml-errors`](../api/xml-errors.md) for the per-codec parse failures.

Every one of them lives in the `@shbernal/ts-xlsx/errors` subpath and nowhere else, because a
container-level failure belongs to no single codec. A service that only classifies failures
can import that subpath alone and load 4 KB instead of a parser.

Next: [in the browser](./browser.md).
