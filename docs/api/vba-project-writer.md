# Vba Project Writer

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `VbaModuleSource`

<sub>interface</sub>

A module to author. `document` and `designer` kinds are host-coupled and not yet synthesizable.

```ts
interface VbaModuleSource {
    readonly name: string;
    readonly kind: 'procedural' | 'class';
    readonly source: string;
}
```

---

### `VbaProjectSpec`

<sub>interface</sub>

```ts
interface VbaProjectSpec {
    readonly modules: readonly VbaModuleSource[];
    readonly codePage?: number;
    readonly projectName?: string;
}
```

---

### `writeVbaProject`

<sub>function</sub>

Synthesize a `vbaProject.bin` from `spec`. The returned bytes can be attached to a workbook via
`Workbook.vbaProjectBytes` (or re-parsed with `parseVbaProject`); the written workbook opens in Excel
as a macro-enabled book whose modules recompile from the embedded source.

```ts
function writeVbaProject(spec: VbaProjectSpec): Uint8Array;
```

**Throws** — on an invalid or duplicate module name, an unsupported module kind, or source
/ a name that the chosen code page cannot represent.
