# Preserved

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `PreservedPart`

<sub>interface</sub>

A package part the model does not interpret, captured verbatim so a round-trip re-emits it intact.
`bytes` are the raw part contents, `contentType` how the source package declared it, and `rels` its
outbound relationships (empty when the part references nothing). The writer re-numbers the part to a
fresh, collision-proof path and rewires `rels` accordingly, but never touches `bytes`.

```ts
interface PreservedPart {
    readonly path: string;
    readonly contentType: string;
    readonly bytes: Uint8Array;
    readonly rels: readonly PreservedRelationship[];
}
```

---

### `PreservedRelationship`

<sub>interface</sub>

One outbound relationship of a `PreservedPart`: the id it carries inside its own rels part,
the relationship Type URI, and its target. An internal relationship's `targetPath` is the resolved
package path of the part it points at (the writer re-numbers and rewires it); an `external`
relationship's `targetPath` is the raw `Target` verbatim (a linked workbook's path or URL) — it is
outside the package, so it is emitted unchanged with `TargetMode="External"` and never remapped.
Preserving external relationships is what keeps an `externalLink` part's pointer to its source
workbook alive, so a round-trip does not orphan the `[n]` external references formulas resolve through.

```ts
interface PreservedRelationship {
    readonly id: string;
    readonly type: string;
    readonly targetPath: string;
    readonly external?: boolean;
}
```

---

### `PreservedRootReference`

<sub>interface</sub>

A package-root reference to content the model does not model, wired from the package's own
`_rels/.rels` rather than the workbook part's rels — the ribbon-customisation parts
(`customUI/customUI14.xml`), custom document properties (`docProps/custom.xml`), a thumbnail, and
anything else hung off the root. The writer regenerates the root rels for the parts it models
(workbook, core/app properties), so these would be dropped unless captured here and re-declared.
`relType` is the relationship Type URI to re-emit, `entryPath` the part it targets (kept at its
original path, since the writer generates nothing of these kinds to collide with), and `parts` the
transitive closure the reference reaches.

```ts
interface PreservedRootReference {
    readonly relType: string;
    readonly entryPath: string;
    readonly parts: readonly PreservedPart[];
}
```

---

### `PreservedWorksheetReference`

<sub>interface</sub>

A worksheet-level reference to package content the model does not model — preserved verbatim across
a round-trip instead of being silently dropped. `element` is the worksheet child that wires the
reference (`<drawing>` for a vector-shape drawing, `<legacyDrawingHF>` for a header/footer image),
or `undefined` when the sheet wires it by relationship alone (a pivot table or slicer Excel
discovers by scanning the sheet's rels, with no worksheet child pointing at it). `relType` is the
relationship Type URI to re-emit; `entryPath` is the part it points at; `parts` is the transitive
closure of parts that reference reaches (the entry included), each re-emitted with its relationships
rewired.

```ts
interface PreservedWorksheetReference {
    readonly element: 'drawing' | 'legacyDrawingHF' | undefined;
    readonly relType: string;
    readonly entryPath: string;
    readonly parts: readonly PreservedPart[];
}
```
