# Xlsx Edit Vba

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `editXlsxVbaAddReference`

<sub>function</sub>

Add a registered (COM type-library) reference to an existing macro-enabled package's VBA project,
returning new package bytes. Every part but `xl/vbaProject.bin` is preserved byte-for-byte (see
`addVbaReference` for what changes within it), and any digital signature over the old project is
dropped because it cannot validate the new bytes.

```ts
function editXlsxVbaAddReference(xlsx: Uint8Array, ref: VbaLibraryReference): Uint8Array;
```

**Throws** — `VbaAuthorError` if the package carries no VBA project, or any field of `ref` is invalid (see
`VbaLibraryReference`).
**Throws** — `VbaParseError` if the attached `vbaProject.bin` is malformed.

---

### `editXlsxVbaRemoveModule`

<sub>function</sub>

Remove a standard module from an existing macro-enabled package's VBA project, returning new package
bytes. Every part but `xl/vbaProject.bin` is preserved byte-for-byte (see `removeVbaModule` for
what changes within it), and any digital signature over the old project is dropped because it cannot
validate the new bytes.

```ts
function editXlsxVbaRemoveModule(xlsx: Uint8Array, name: string): Uint8Array;
```

**Throws** — `VbaAuthorError` if the package carries no VBA project, `name` is not in the project, or names
a `document`/`designer` module.
**Throws** — `VbaParseError` if the attached `vbaProject.bin` is malformed.
