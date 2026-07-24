# Reading .xlsx

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `readXlsx`

<sub>function</sub>

Read an `.xlsx` package into a `Workbook`.

```ts
function readXlsx(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook;
```

**Throws** — if the input is not a readable `.xlsx` package — a legacy `.xls`
(`.format === 'xls'`), a binary `.xlsb` (`'xlsb'`), or an unrecognised/non-ZIP blob (`'unknown'`).
**Throws** — if the archive exceeds the inflate bound (a probable zip bomb).

---

### `ReadXlsxOptions`

<sub>interface</sub>

```ts
interface ReadXlsxOptions {
    readonly maxUncompressedBytes?: number;
}
```
