# Customui Ribbon

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CUSTOMUI_2006_NAMESPACE`

<sub>const</sub>

The `customUI` root namespaces, one per `RibbonDialect`.

```ts
const CUSTOMUI_2006_NAMESPACE: "http://schemas.microsoft.com/office/2006/01/customui"
```

---

### `CUSTOMUI_2007_REL_TYPE`

<sub>const</sub>

```ts
const CUSTOMUI_2007_REL_TYPE: "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
```

---

### `CUSTOMUI_2009_NAMESPACE`

<sub>const</sub>

```ts
const CUSTOMUI_2009_NAMESPACE: "http://schemas.microsoft.com/office/2009/07/customui"
```

---

### `CUSTOMUI_2010_REL_TYPE`

<sub>const</sub>

```ts
const CUSTOMUI_2010_REL_TYPE: "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
```

---

### `CustomUiDocument`

<sub>interface</sub>

A parsed `customUI` part. `dialect` records which schema it was written against (derived from the
root namespace, the authoritative signal). `ribbon` is the parsed `<ribbon>` subtree, or `undefined`
when the document customises only backstage/QAT/commands (which v1 does not parse). Future work can
extend this with `backstage`/`qat` without changing the shape callers already depend on.

```ts
interface CustomUiDocument {
    readonly dialect: RibbonDialect;
    readonly ribbon?: Ribbon;
}
```

---

### `isCustomUiRelType`

<sub>function</sub>

Whether a package-root relationship Type URI points at a `customUI` ribbon part.

```ts
function isCustomUiRelType(type: string): boolean;
```

---

### `parseCustomUi`

<sub>function</sub>

Parse a `customUI` part (raw UTF-8 bytes or its decoded text) into a `CustomUiDocument`.

```ts
function parseCustomUi(input: string | Uint8Array): CustomUiDocument;
```

**Throws** — {

---

### `Ribbon`

<sub>interface</sub>

The parsed `<ribbon>` element: whether it starts from a blank ribbon, and its custom tabs. Only the
`<tabs>` subtree is modelled; `qat` and `contextualTabs` are not parsed in v1.

```ts
interface Ribbon {
    readonly startFromScratch: boolean;
    readonly tabs: readonly RibbonTab[];
}
```

---

### `RibbonControl`

<sub>interface</sub>

A control element inside a ribbon group. `kind` is the element's local name, narrowed to the closed
set of RibbonX control elements (`RibbonControlKind`); an element outside that set is surfaced
as `unknown` rather than dropped. The three identity attributes (`id` a document-defined control,
`idQ` a qualified id, `idMso` a built-in control) and the two most-consulted display/behaviour
attributes (`label`, `onAction`) are lifted out as typed conveniences; every attribute the element
actually carried — including the many `get*` dynamic callbacks and layout hints not modelled here —
is preserved verbatim in `attributes`, so nothing is lost. Container controls (a `menu`,
`splitButton`, `gallery`, `dropDown`, `box`, …) carry their nested controls/items in `children`.

```ts
interface RibbonControl {
    readonly kind: RibbonControlKind;
    readonly id?: string;
    readonly idQ?: string;
    readonly idMso?: string;
    readonly label?: string;
    readonly onAction?: string;
    readonly attributes: Readonly<Record<string, string>>;
    readonly children?: readonly RibbonControl[];
}
```

---

### `RibbonControlKind`

<sub>type</sub>

The RibbonX control elements this reader recognises. `item` is a `dropDown`/`gallery`/`comboBox`
entry; `unknown` is the fallback for any element outside this set (never silently dropped).

```ts
type RibbonControlKind = 'button' | 'toggleButton' | 'checkBox' | 'editBox' | 'dropDown' | 'comboBox' | 'gallery' | 'menu' | 'dynamicMenu' | 'splitButton' | 'buttonGroup' | 'box' | 'labelControl' | 'separator' | 'menuSeparator' | 'dialogBoxLauncher' | 'control' | 'item' | 'unknown';
```

---

### `RibbonDialect`

<sub>type</sub>

The `customUI` schema a part is written against — the read model keys off this, not the (frequently
mis-copied) relationship type. `2007` is the original RibbonX (`customUI.xml`); `2010` is the later
schema (`customUI14.xml`) that also carries backstage/QAT/commands.

```ts
type RibbonDialect = '2007' | '2010';
```

---

### `RibbonGroup`

<sub>interface</sub>

A `<group>` within a ribbon tab: its identity/label attributes and the controls it contains.

```ts
interface RibbonGroup {
    readonly id?: string;
    readonly idQ?: string;
    readonly idMso?: string;
    readonly label?: string;
    readonly attributes: Readonly<Record<string, string>>;
    readonly controls: readonly RibbonControl[];
}
```

---

### `RibbonTab`

<sub>interface</sub>

A `<tab>` within the ribbon: its identity/label attributes and the groups it contains.

```ts
interface RibbonTab {
    readonly id?: string;
    readonly idQ?: string;
    readonly idMso?: string;
    readonly label?: string;
    readonly attributes: Readonly<Record<string, string>>;
    readonly groups: readonly RibbonGroup[];
}
```
