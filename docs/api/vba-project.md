# Vba Project

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `parseVbaProject`

<sub>function</sub>

```ts
function parseVbaProject(bin: Uint8Array): VbaProject;
```

---

### `VbaModule`

<sub>interface</sub>

```ts
interface VbaModule {
    readonly name: string;
    readonly streamName: string;
    readonly kind: VbaModuleKind;
    readonly source: string;
}
```

---

### `VbaModuleKind`

<sub>type</sub>

How a module participates in the project — the classification the VBA editor shows.

```ts
type VbaModuleKind = 'procedural' | 'document' | 'class' | 'designer';
```

---

### `VbaProject`

<sub>interface</sub>

```ts
interface VbaProject {
    readonly codePage: number;
    readonly modules: readonly VbaModule[];
}
```

---

### `VbaProjectSignature`

<sub>interface</sub>

One digital signature over a workbook's VBA project — its generation and its raw signature bytes.

```ts
interface VbaProjectSignature {
    readonly kind: VbaProjectSignatureKind;
    readonly bytes: Uint8Array;
}
```

---

### `VbaProjectSignatureKind`

<sub>type</sub>

Which generation of VBA project signature a part is — Office emits up to three sibling signature
parts off `vbaProject.bin`'s own rels over the same project bytes ([MS-OFFMACRO2]): the original
`legacy` signature, the `agile` (V2) successor, and the `v3` scheme that closes a tampering hole
the earlier two left open (KB5000676). All three can coexist in one package.

```ts
type VbaProjectSignatureKind = 'legacy' | 'agile' | 'v3';
```
