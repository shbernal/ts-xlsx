# Reading .xlsx

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `readXlsx`

<sub>function</sub>

Read an `.xlsx` package into a `Workbook`.

```ts
function readXlsx(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook;
```

**Throws** — if the archive is malformed, exceeds the inflate bound, or names no
worksheet parts (a workbook with no sheets is not a valid package).

---

### `ReadXlsxOptions`

<sub>interface</sub>

```ts
interface ReadXlsxOptions {
    readonly maxUncompressedBytes?: number;
}
```
