# In the browser

The buffered half of this library was built to run in a tab. It is synchronous, it speaks
`Uint8Array` rather than Node's `Buffer`, and reading or writing a workbook is a function
call rather than a stream to await. The
[playground](https://shbernal.github.io/ts-xlsx/playground) is that claim running: it writes,
reads and round-trips workbooks entirely in the page, with no server behind it.

## What runs in a browser

`readXlsx`, `writeXlsx`, `writeXlsxAsync`, the whole `Workbook` model, the CSV functions and
every error class. That is the buffered API, and it is the API most code uses.

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

Handing bytes back to the reader is the mirror of that, and it is the platform's job rather
than the library's:

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
`@shbernal/ts-xlsx/node` and the root specifier does not carry it; a browser build that
imports that subpath anyway links a module whose classes throw by name, which is the honest
answer rather than a bundler error about `node:fs`. The streaming reader is a generator over
Node buffers and is not available in a tab either. If your browser code needs either, it
needs a server.

Everything else runs, including the two things that used to be listed here. Sheet-protection
passwords are hashed by this library's own SHA-512 rather than by `node:crypto`, and the CSV
encoders are the platform's `TextEncoder` and two small loops rather than `Buffer`.

## Nothing to configure

There is no bundler alias to write and no `node:` specifier to stub out. Nothing reachable
from `@shbernal/ts-xlsx`, `/core`, `/xlsx`, `/xlsb`, `/csv`, `/vba`, `/customui` or `/errors`
imports a Node built-in or reads a Node global, and that is a gate rather than a claim:
`scripts/check-browser-safe.ts` walks the module graph from every one of those entries on
every run, and the published build is walked again from the emitted JavaScript. If a Node
import ever reaches the browser entries, CI says so before you do.

The one subpath that does reach them, `/node`, is resolved by a bundler's `browser`
condition to a module that throws with the name and the way out:

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

You can do better than that number if you need to, because the package declares
`"sideEffects": false` and exports subpaths. With a bundler, importing from the package root
is the right default and unused modules are dropped whole. Reach for a subpath when you would
rather the module graph itself said which half of the library you depend on:

```ts
import {Workbook} from '@shbernal/ts-xlsx/core';
import {XlsxError} from '@shbernal/ts-xlsx/errors';

const sheet = new Workbook().addWorksheet('S');
sheet.getCell('A1').value = 'core alone, with no codec behind it';
console.log(XlsxError.name); // 'XlsxError'
```

`/core` is the model with no reader or writer; `/errors` is the failure taxonomy alone, which
is what a service that only classifies failures needs. The full table of entry points and
what each pulls in is in the [README](../../README.md).

## Reading a file someone gave you

Everything in [reading a workbook](./reading.md) about untrusted input applies with more
force here, because in a browser the file arrived from a person rather than from your own
storage. The two defences are on and are not configurable: inflation is bounded by counting
output rather than by trusting declared sizes, and XML entities are decoded but never
expanded. You can lower the inflate ceiling from its 512 MiB default if you know your inputs
are small.

That is the whole of the browser story. Back to the [guide index](./README.md).
