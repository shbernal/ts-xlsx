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
  /** The module's code name as seen in the VBA editor, e.g. `ThisWorkbook`, `JsonConverter`. */
  readonly name: string;
  /** The CFB stream the module's bytes live in — usually equal to {@link name}. */
  readonly streamName: string;
  /** Procedural (`.bas`), document code-behind, class module, or designer (UserForm). */
  readonly kind: VbaModuleKind;
  /** The decompressed VBA source (p-code and PerformanceCache are not included). */
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
  /** The project code page (`PROJECTCODEPAGE`) used to decode module names and source. */
  readonly codePage: number;
  /** The project's modules, in declaration order. */
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
  /**
   * The raw signature part bytes (a PKCS#7/CMS blob), passed through verbatim — this library does not
   * parse or cryptographically verify them. Their presence means "a signature is attached," never
   * "this signature is valid."
   */
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
