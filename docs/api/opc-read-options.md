# Opc Read Options

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `ReadXlsxOptions`

<sub>interface</sub>

```ts
interface ReadXlsxOptions {
  /**
   * Maximum total uncompressed output, in bytes, produced while inflating the package.
   * The bound is enforced by a running counter as bytes are decompressed — never read from
   * the archive's (untrusted, forgeable) size headers — so a zip bomb that lies about its
   * uncompressed size is rejected all the same. Defaults to 512 MiB.
   */
  readonly maxUncompressedBytes?: number;
}
```
