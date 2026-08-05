# Writing .xlsx

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `WriteOptions`

<sub>interface</sub>

Options controlling how `writeXlsx` serialises a workbook.

```ts
interface WriteOptions {
  /**
   * Pool plain string cell values into a shared-strings table (`xl/sharedStrings.xml`) that cells
   * reference by index, rather than storing each string inline in its cell. Deduplicates repeated
   * text and matches Excel's own storage; off by default, which keeps strings inline and omits the
   * part. Rich-text values stay inline regardless, so their run formatting is unaffected.
   */
  readonly useSharedStrings?: boolean;
}
```

---

### `writeXlsx`

<sub>function</sub>

Serialise a workbook into an `.xlsx` package.

```ts
function writeXlsx(workbook: Workbook, options: WriteOptions = {}): Uint8Array;
```

**Throws** — {

---

### `writeXlsxAsync`

<sub>function</sub>

Serialise a workbook into an `.xlsx` package, deflating off the calling thread.

Produces the same package `writeXlsx` does — every part compresses to identical bytes — and
exists for one reason: DEFLATE dominates the cost of writing a large workbook, and `writeXlsx`
spends all of it on the caller's thread. Here `fflate` deflates each part in a worker, so the event
loop keeps turning (stalls drop from the whole write to tens of milliseconds) and parts compress in
parallel, which on a multi-sheet workbook also finishes sooner. On a single-sheet workbook there is
only one part to deflate, so expect responsiveness rather than speed.

Building the parts still happens on the calling thread — only compression moves. That is why there
is no `readXlsxAsync` mirroring this: reading is dominated by XML parsing and model building, which
no worker can take, and the reader's zip-bomb ceiling is enforced by counting output between
synchronous input slices. See ADR-0024.

```ts
async function writeXlsxAsync(
  workbook: Workbook,
  options: WriteOptions = {},
): Promise<Uint8Array>;
```

**Throws** — {
