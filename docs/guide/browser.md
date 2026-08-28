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

The streaming writer writes to a Node stream and the streaming reader is a Node-shaped
generator over Node buffers; neither is available in a tab. Sheet-protection passwords are
hashed with Node's crypto. If your browser code needs any of those, it needs a server.

## The gap you will hit today, and what to do about it

Three Node built-ins are statically reachable from the package entry: `node:crypto` through
the sheet-protection password hash, and `node:fs` and `node:stream` through the streaming
writer. Nothing on the buffered path calls any of them, but a bundler resolves imports rather
than call graphs, so it will reach all three anyway. What you see depends on the bundler:
Vite externalises them with a warning and then fails the build on a named import; webpack
reports a module it cannot resolve.

This is a real limitation and not a configuration mistake on your side. Until the library
puts those behind an export condition a browser never resolves, the fix is to alias the three
specifiers to a module that throws:

```ts
// node-absent.ts
function absent(name: string): never {
  throw new Error(`${name} is not available in a browser; use readXlsx and writeXlsx.`);
}

export const createHash = (): never => absent('createHash');
export const randomBytes = (): never => absent('randomBytes');
export const createWriteStream = (): never => absent('createWriteStream');
export class PassThrough {
  constructor() {
    absent('PassThrough');
  }
}

console.log(typeof createHash); // 'function'
```

Then point the bundler at it. In Vite:

<!-- sample: illustrative. It is a fragment of a build config rather than a program. -->

```ts
export default {
  resolve: {
    alias: [{find: /^node:(?:crypto|fs|stream)$/, replacement: '/src/node-absent.ts'}],
  },
};
```

This site does exactly that, which is why the playground works. The errors it throws never
fire in practice, because nothing the buffered path does reaches them.

## What it costs to ship

The site's own production build puts the whole library, `fflate`, and the playground's page
code into one chunk of 331 KB, which is 92 KB over the wire once gzipped. That chunk is
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
