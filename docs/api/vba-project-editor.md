# Vba Project Editor

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `editVbaModuleSources`

<sub>function</sub>

Edit the source of one or more existing modules in a `vbaProject.bin`, returning new bytes that
preserve the project's references, host info, and every other module. `edits` maps a module's code
name (case-insensitively, as VBA compares them) to its replacement source.

```ts
function editVbaModuleSources(bin: Uint8Array, edits: ReadonlyMap<string, string>): Uint8Array;
```

**Throws** — if `bin` is not a parseable VBA project (validated before any edit).
**Throws** — if a named module is absent, or the new source has a character the project's
code page cannot represent.
