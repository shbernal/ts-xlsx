# In the browser

The buffered half of this library runs in a browser tab. It is synchronous, it uses
`Uint8Array` instead of Node's `Buffer`, and reading or writing a workbook is one function
call. The [playground](https://shbernal.github.io/ts-xlsx/playground) writes, reads and
round-trips workbooks entirely in the page, with no server behind it.

## What runs in a browser

`readXlsx`, `writeXlsx`, `writeXlsxAsync`, the whole `Workbook` model, the CSV functions and
every error class. That is the buffered API, which is what most code uses.

<!-- sample: illustrative. It needs a File from an input element, which a guide has none of. -->

```ts
import {readXlsx} from '@shbernal/ts-xlsx';

input.addEventListener('change', async () => {
  const file = input.files?.[0];
  if (file === undefined) return;
  const workbook = readXlsx(new Uint8Array(await file.arrayBuffer()));
  console.log(workbook.worksheets.map((sheet) => sheet.name));
});
```

Saving the bytes as a download is the platform's job, not the library's:

<!-- sample: illustrative. It needs a document to create the link in. -->

```ts
import {writeXlsx} from '@shbernal/ts-xlsx';

const blob = new Blob([writeXlsx(workbook)], {
  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
});
const url = URL.createObjectURL(blob);
// Give the url to an <a download>, then revoke it once the click has happened.
URL.revokeObjectURL(url);
```

`writeXlsxAsync` hands compression to fflate's asynchronous deflate, which uses workers where
the environment provides them. It produces the same package byte for byte.

## What does not

The streaming writer. It opens files and pipes Node streams, so it is published from
`@shbernal/ts-xlsx/node` and the root specifier does not carry it. A browser build that
imports that subpath anyway links a module whose classes throw an error that names the
problem, which beats a bundler error about `node:fs`. The streaming reader is a generator over
Node buffers and is not available in a tab either. If your browser code needs either, it
needs a server.

Everything else runs. Sheet-protection passwords are hashed with the library's own SHA-512
instead of `node:crypto`, and the CSV encoders use the platform's `TextEncoder` instead of
`Buffer`.

## Nothing to configure

There is no bundler alias to write and no `node:` specifier to stub out. Nothing reachable
from `@shbernal/ts-xlsx`, `/core`, `/xlsx`, `/xlsb`, `/csv`, `/vba`, `/customui` or `/errors`
imports a Node built-in or reads a Node global. A CI check enforces that:
`scripts/check-browser-safe.ts` walks the module graph from every one of those entries on
every run, then walks the published build again from the emitted JavaScript. If a Node
import ever reaches a browser entry, CI fails.

The one subpath that does use them, `/node`, resolves under a bundler's `browser`
condition to a module that throws an error naming the problem and the fix:

<!-- sample: illustrative. It shows what a browser build links, which this page is not. -->

```ts
import {WorkbookStreamWriter} from '@shbernal/ts-xlsx/node';

// In a browser build:
new WorkbookStreamWriter();
// Error: ts-xlsx: WorkbookStreamWriter is not available in this environment: the streaming
// writer opens files and pipes Node streams, which a browser has neither of. Use writeXlsx
// (or writeXlsxAsync), which produce the same package as bytes.
```

## What it costs to ship

The site's own production build puts the whole library, `fflate`, and the playground's page
code into one chunk of 328 KB, which is 93 KB over the wire once gzipped. That chunk is
fetched when the playground mounts and by no other page.

You can ship less, because the package declares `"sideEffects": false` and exports subpaths.
With a bundler, import from the package root and let it drop unused modules. Use a subpath
when you want the module graph itself to say which part of the library you depend on:

```ts
import {Workbook} from '@shbernal/ts-xlsx/core';
import {XlsxError} from '@shbernal/ts-xlsx/errors';

const sheet = new Workbook().addWorksheet('S');
sheet.getCell('A1').value = 'core alone, with no codec behind it';
console.log(XlsxError.name); // 'XlsxError'
```

`/core` is the model with no reader or writer. `/errors` is the error classes alone, which
is all a service that only classifies failures needs.

| Import from                  | You get                                                                        | It loads |
| ---------------------------- | ------------------------------------------------------------------------------ | -------- |
| `@shbernal/ts-xlsx`          | everything except the streaming writer                                         | 670 KB   |
| `@shbernal/ts-xlsx/core`     | `Workbook`, `Worksheet`, `Cell`, styles, values, addresses                     | 245 KB   |
| `@shbernal/ts-xlsx/xlsx`     | `readXlsx`, `writeXlsx`/`writeXlsxAsync`, the streaming reader, VBA part edits | 652 KB   |
| `@shbernal/ts-xlsx/xlsb`     | `readXlsb`                                                                     | 357 KB   |
| `@shbernal/ts-xlsx/csv`      | `readCsv`, `writeCsv`, `writeCsvText`                                          | 256 KB   |
| `@shbernal/ts-xlsx/node`     | `WorkbookStreamWriter` and the rest of the streaming writer (Node only)        | 500 KB   |
| `@shbernal/ts-xlsx/vba`      | `parseVbaProject`, `addVbaReference`, `removeVbaModule`                        | 55 KB    |
| `@shbernal/ts-xlsx/customui` | `parseCustomUi` and the ribbon types                                           | 15 KB    |
| `@shbernal/ts-xlsx/errors`   | every error class the library throws                                           | 4 KB     |

The sizes are the unminified static-import closure of each entry, measured by
`pnpm run size` at 4.0.0. A bundler and minifier only shrink them.

`/xlsx` is barely smaller than the whole package. `readXlsx` sniffs the bytes and hands a
binary package to the `.xlsb` reader, so that codec is always on its path. Every error class
lives only in `/errors`, because a failure like `UnsupportedFormatError` belongs to no single
codec: `readXlsx` and `readXlsb` both raise it.
[ADR-0040](../decisions/0040-the-browser-boundary-is-an-entry-point.md) records why `/node`
is the one entry the root does not carry.

## Reading a file someone gave you

Everything in [reading a workbook](./reading.md) about untrusted input applies even more
here, because in a browser the file came from a person, not from your own storage. The two defences are on and are not configurable: inflation is bounded by counting
output rather than by trusting declared sizes, and XML entities are decoded but never
expanded. You can lower the inflate ceiling from its 512 MiB default if you know your inputs
are small.

Back to the [guide index](./README.md).
