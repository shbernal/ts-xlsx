# Opc Read Options

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `ReadPackageOptions`

<sub>interface</sub>

How a reader is allowed to open an OPC package. One knob today: the ceiling on what inflating it
may produce, which every reader shares because every reader inflates the same container.

```ts
interface ReadPackageOptions {
  /**
   * Maximum total uncompressed output, in bytes, produced while inflating the package.
   * The bound is enforced by a running counter as bytes are decompressed, never read from
   * the archive's (untrusted, forgeable) size headers, so a zip bomb that lies about its
   * uncompressed size is rejected all the same. Defaults to 512 MiB.
   */
  readonly maxUncompressedBytes?: number;
}
```
