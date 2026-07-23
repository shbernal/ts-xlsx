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
