# Xlsx Edit Vba

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `editXlsxVbaModuleSource`

<sub>function</sub>

Edit one existing VBA module's source in a macro-enabled package, preserving every other part exactly.
A convenience over `editXlsxVbaModuleSources` for the common single-module case.

```ts
function editXlsxVbaModuleSource(xlsx: Uint8Array, name: string, source: string): Uint8Array;
```

**Throws** — if the package carries no VBA project, the module is absent, or the new
source has a character the project's code page cannot represent.
**Throws** — if the attached `vbaProject.bin` is malformed.

---

### `editXlsxVbaModuleSources`

<sub>function</sub>

Edit the source of one or more existing VBA modules in a macro-enabled `.xlsm` package, returning new
package bytes. Every part but `xl/vbaProject.bin` is preserved byte-for-byte; the macro project is
spliced in place (references, host info, and untouched modules kept — see
`editVbaModuleSources`), and any digital signature over the old project is dropped because it
cannot validate the new bytes. `edits` maps a module's code name (case-insensitively, as VBA compares
them) to its replacement source.

Unlike a `readXlsx`/`writeXlsx` round-trip, this does not re-serialise the workbook from a model, so
the non-macro content of a real-world file survives exactly.

```ts
function editXlsxVbaModuleSources(xlsx: Uint8Array, edits: ReadonlyMap<string, string>): Uint8Array;
```

**Throws** — if the package carries no VBA project, a named module is absent, or a new
source has a character the project's code page cannot represent.
**Throws** — if the attached `vbaProject.bin` is malformed.
