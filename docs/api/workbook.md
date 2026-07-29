# Workbook

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `AddImageOptions`

<sub>interface</sub>

A picture registered on the workbook, ready to be anchored to a worksheet.

```ts
interface AddImageOptions {
    readonly buffer: Uint8Array;
    readonly extension?: string;
}
```

---

### `AddWorksheetOptions`

<sub>interface</sub>

```ts
interface AddWorksheetOptions {
    readonly state?: WorksheetState['state'];
}
```

---

### `DEFAULT_WORKBOOK_VIEW`

<sub>const</sub>

The window geometry a workbook starts from — the values desktop Excel writes for its own default
window.

A default is emitted rather than the element left out because Excel writes `<bookViews>` into every
file it saves and consumers lay panes out against that rect. With no view at all the frozen-pane
split is computed against an uninitialised window, and the frozen region can stay unpainted until
some later event forces a relayout.

```ts
const DEFAULT_WORKBOOK_VIEW: { readonly x: -110; readonly y: -110; readonly width: 19420; readonly height: 12220; readonly activeTab: 0; }
```

---

### `DefinedName`

<sub>interface</sub>

A named reference in the workbook — the entries Excel surfaces in its Name Manager. A name maps
an identifier to a formula (`refersTo`), most often a cell range like `Sheet1!$A$1:$B$2` but
possibly any formula. A name is global to the workbook unless it names a sheet in `scope`,
which restricts it to that sheet and lets another sheet reuse the same name independently.

```ts
interface DefinedName {
    readonly name: string;
    readonly refersTo: string;
    readonly scope?: string;
    readonly comment?: string;
    readonly hidden?: boolean;
}
```

---

### `PreservedWorkbookReference`

<sub>interface</sub>

A workbook-level reference to package content the model does not model — a pivot cache
(`pivotCacheDefinition`) or a slicer cache (`slicerCache`) — preserved verbatim across a round-trip
instead of being dropped. `relType` is the workbook relationship Type URI to re-emit; `entryPath`
is the part it points at; `parts` is the transitive closure that reference reaches (the entry
included). `pivotCacheId` carries the `<pivotCache cacheId>` a pivot cache is registered under in
the workbook's `<pivotCaches>`, so the wiring a pivot table resolves its cache through is re-emitted
too; it is absent for a slicer cache, which the workbook lists in an extension block instead.
`externalReferenceIndex` is the 0-based position of an `externalLink` within the workbook's
`<externalReferences>` — the `[n]` a formula or defined name resolves an external cell through — so
the block is re-emitted in the original order and every `[n]` still points at the same linked
workbook; it is absent for a pivot/slicer cache.

```ts
interface PreservedWorkbookReference {
    readonly relType: string;
    readonly entryPath: string;
    readonly parts: readonly PreservedPart[];
    readonly pivotCacheId?: string;
    readonly externalReferenceIndex?: number;
}
```

---

### `Workbook`

<sub>class</sub>

```ts
class Workbook {
  readonly properties: WorkbookProperties = {};
  readonly view: WorkbookView = { ...DEFAULT_WORKBOOK_VIEW };
  fullCalcOnLoad = false;
  protection: WorkbookProtection | undefined = undefined;
  get worksheets(): readonly Worksheet[];
  get activeTabIndex(): number;
  get preservedReferences(): readonly PreservedWorkbookReference[];
  get preservedRootReferences(): readonly PreservedRootReference[];
  get customUI(): readonly CustomUiDocument[];
  get vbaProject(): VbaProject | undefined;
  get vbaProjectBytes(): Uint8Array | undefined;
  set vbaProjectBytes(bytes: Uint8Array | undefined);
  get vbaProjectSigned(): boolean;
  get vbaProjectSignatures(): readonly VbaProjectSignature[];
  removeVbaModule(name: string): void;
  addVbaReference(ref: VbaLibraryReference): void;
  get differentialStyles(): readonly string[];
  get indexedColors(): readonly string[];
  get mruColors(): readonly string[];
  get tableStyles(): TableStyleTable;
  addTableStyle(style: TableStyle): void;
  get customTableStyles(): readonly TableStyle[];
  get themePart(): PreservedTheme | undefined;
  setTheme(overrides: ThemeOverrides): void;
  get themeColors(): ThemeColorScheme;
  get themeFonts(): ThemeFontScheme;
  get declaredDefaultFont(): Font | undefined;
  setDefaultFont(font: Font): void;
  get defaultFont(): Font;
  authoredThemeXml(): string | undefined;
  resolveColor(color: Color): string | undefined;
  get namedStyles(): readonly NamedCellStyle[];
  addPerson(person: Person): void;
  get persons(): readonly Person[];
  getPerson(id: string): Person | undefined;
  addImage(options: AddImageOptions): number;
  get media(): readonly WorkbookImage[];
  getImage(id: number): WorkbookImage | undefined;
  get definedNames(): readonly DefinedName[];
  defineName(definedName: DefinedName): void;
  addWorksheet(name: string, options: AddWorksheetOptions = {}): Worksheet;
  getWorksheet(nameOrId: string | number): Worksheet | undefined;
}
```

**Members**

- `readonly view: WorkbookView = { ...DEFAULT_WORKBOOK_VIEW };` — The workbook's window state — position, size, and the selected sheet. Always present (see `DEFAULT_WORKBOOK_VIEW` for why it is defaulted rather than left unset) and always written. Reading a file replaces it with that file's saved geometry, so a round-trip restores the window the author left rather than stamping ours over it.
- `fullCalcOnLoad = false;` — Ask consuming spreadsheet apps to recalculate every formula when the file is opened, rather than trusting the cached results stored with each formula cell. Set this when the producer cannot compute formula results itself — the OOXML `fullCalcOnLoad` flag. Off by default, so a workbook whose cached results are authoritative stays unmarked.
- `protection: WorkbookProtection | undefined = undefined;` — Workbook-level structure/window protection — the OOXML `<workbookProtection>` element. Absent by default (an unprotected workbook). Set it to lock the workbook shell, or leave it as read from a file so a protected workbook stays locked across a passthrough save rather than being silently unlocked. Distinct from a worksheet's own `protect()`, which guards a single sheet's cells.
- `get worksheets(): readonly Worksheet[];` — The worksheets in insertion order.
- `get activeTabIndex(): number;` — The 0-based index of the active sheet: `WorkbookView.activeTab` resolved against the sheets that actually exist. Exactly one sheet is always active — an out-of-range tab (a caller's stale index, or a file whose sheet was removed after the view was saved) falls back to the first sheet rather than to none, because a package where no sheet is selected gives the consumer no view to initialise on open.
- `get preservedReferences(): readonly PreservedWorkbookReference[];` — The workbook-level preserved references, in the order they were read.
- `get preservedRootReferences(): readonly PreservedRootReference[];` — The package-root preserved references, in the order they were read.
- `get customUI(): readonly CustomUiDocument[];` — The ribbon customisations decoded from this workbook's `customUI` parts — `customUI.xml` (Office 2007) and/or `customUI14.xml` (Office 2010+), in the order their root relationships were read. Each `CustomUiDocument` is tagged with its dialect and exposes the parsed `<ribbon>` tree. Empty for a workbook that customises no ribbon. This is a **read-only view** over parts the writer already round-trips verbatim — mutating the returned objects changes nothing on write; the original `customUI` XML is re-emitted byte-for-byte regardless. Parsed lazily on first access and memoised.
- `get vbaProject(): VbaProject | undefined;` — The VBA project decoded from this workbook's preserved `vbaProject.bin`, or `undefined` for a workbook with no macros. This is a **read-only view** over the bytes the writer already round-trips verbatim — mutating the returned object changes nothing on write; the original macro blob is re-emitted byte-for-byte regardless. Parsed lazily on first access and memoised.
- `get vbaProjectBytes(): Uint8Array | undefined;` — The raw `vbaProject.bin` bytes attached to this workbook — the exact macro blob the writer will embed — or `undefined` for a workbook with no macros. The getter returns a defensive copy, so mutating it changes nothing on write. Assigning bytes attaches (or replaces) the macro project: the written package becomes macro-enabled and re-embeds these bytes verbatim. The bytes must be a well-formed VBA container (a CFB holding a `dir` stream); a malformed blob is rejected with `VbaParseError` rather than written out to produce a package Excel would flag for repair. This is the attach-blob path: copy a project between workbooks with `dst.vbaProjectBytes = src.vbaProjectBytes`, or import a `.bin` produced by another tool. Assigning `undefined` removes the project, reverting the workbook to a plain (non-macro) package. Replacing or removing the project also drops any digital signature the previous blob carried — a signature over the old bytes cannot validate new ones — so the result never advertises a broken signature.
- `get vbaProjectSigned(): boolean;` — Whether this workbook's VBA project carries a digital signature — `true` if any signature part is attached, `false` for an unsigned project or a workbook with no macros. This reflects the **presence** of a signature blob, not its cryptographic validity: the library neither parses the PKCS#7/CMS structure nor validates the certificate chain or signer. A `true` here means "a signature is attached," never "this signature is valid." Replacing or editing the project drops its signatures (a signature over the old bytes cannot validate new ones), so this reads `false` again after `vbaProjectBytes`, `removeVbaModule`, or `addVbaReference` mutates the project. See `vbaProjectSignatures` for the raw bytes and which generation(s) are present.
- `get vbaProjectSignatures(): readonly VbaProjectSignature[];` — The digital signatures attached to this workbook's VBA project, in the order their relationships are wired off `vbaProject.bin` — up to three generations (legacy, agile, V3) can coexist over the same project bytes. Empty for an unsigned project or a workbook with no macros. Each entry's `bytes` are the raw signature blob passed through verbatim; the library does not parse or verify them (see `vbaProjectSigned` on presence-vs-validity). Hand a blob to an external verifier if you need cryptographic validation — that is deliberately out of this library's scope.
- `removeVbaModule(name: string): void;` — Remove a standard module from this workbook's existing macro project, in place — a structural splice that leaves every remaining module's compiled p-code untouched (see `removeVbaModule`). Replacing the project also drops a stale signature, as `vbaProjectBytes` does. Only `procedural` and `class` modules can be removed this way — see `removeVbaModule` for why. To author or edit module *source* (which needs real compiled p-code), use the offline `tools/vba-compiler`, then attach its output via `vbaProjectBytes`.
- `addVbaReference(ref: VbaLibraryReference): void;` — Add a registered (COM type-library) reference to this workbook's existing macro project, in place. Every existing module, reference, and host-info record rides through unchanged (see `addVbaReference`). Replacing the project also drops a stale signature, as `vbaProjectBytes` does.
- `get differentialStyles(): readonly string[];` — The preserved differential-style (`<dxfs>`) fragments, in index order.
- `get indexedColors(): readonly string[];` — The preserved custom indexed-color palette, in index order; empty when the default palette rules.
- `get mruColors(): readonly string[];` — The preserved most-recently-used colour swatches, in order; empty when the file declared none.
- `get tableStyles(): TableStyleTable;` — The preserved `<tableStyles>` block; `styles` is empty when the file declared no custom style.
- `addTableStyle(style: TableStyle): void;` — Register a custom table style — a named look a table applies to itself by putting that name in `TableStyleInfo.name`, exactly as it would name one of Excel's built-in gallery styles. ```ts workbook.addTableStyle({ name: 'Harbour', elements: { wholeTable: {border: {top: {style: 'thin'}, bottom: {style: 'thin'}}}, headerRow: {font: {bold: true, color: {argb: 'FFFFFFFF'}}, fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFBB2649'}}}, firstRowStripe: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFF6E7EB'}}}, }, }); sheet.addTable({name: 'Cargo', ref: 'A1:B3', columns, style: {name: 'Harbour'}}); ``` Each element's formatting is interned into the workbook's shared differential-style table, so two elements — or a conditional-formatting rule — that paint the same way share one entry. Registering a name a source file already defined **overrides** that definition rather than adding a second one beside it.
- `get customTableStyles(): readonly TableStyle[];` — The table styles authored on this workbook, in registration order.
- `get themePart(): PreservedTheme | undefined;` — The preserved theme part, or undefined when the workbook rides the library's default theme.
- `setTheme(overrides: ThemeOverrides): void;` — Author the workbook's theme: any subset of the twelve colour-scheme slots, and either of the two typefaces. Merges into what the workbook already has, so branding one accent leaves the other eleven slots alone, and calling it twice accumulates. This is the workbook-wide palette. A cell that names a colour as `theme="4"` — which is what Excel writes whenever a user picks from the theme row of the colour picker — follows `accent1` here, so one call restyles every such cell, chart and table style at once. Colours are `RRGGBB`; a leading `#` and an 8-hex ARGB are both accepted and reduced, and anything else throws rather than writing a value Excel silently renders as flat black. What it does **not** touch: the theme's format scheme — the gradient, line and effect styles that give a theme its texture. Those ride through from the source theme (or the library's default) untouched, because nobody hand-authors gradient stops from a spreadsheet API and regenerating them would replace a designer's work with the Office default. For the same reason a slot left unauthored keeps the source's own encoding, including the `<a:sysClr>` form Excel uses for `dk1`/`lt1` so they follow the viewer's window colours.
- `get themeColors(): ThemeColorScheme;` — The colour scheme every `theme="n"` reference in this workbook resolves against — anything `setTheme` authored, over the preserved theme's `<a:clrScheme>`, over the Office default. Note the slot *order*: `theme="0"` is `lt1` and `theme="1"` is `dk1`, which is not the order the slots appear in the theme part. See `THEME_COLOR_SLOTS`.
- `get themeFonts(): ThemeFontScheme;` — The theme's major (heading) and minor (body) typefaces, authored values over the source's.
- `get declaredDefaultFont(): Font | undefined;` — The default font as the source package declared it — font id 0 of its styles part, the face every cell that names no font of its own renders in. `undefined` for a workbook authored from scratch or read from a package carrying no styles part: nothing was declared, and the library does not fabricate a declaration on the file's behalf. This is the *round-trip* surface. `defaultFont` is what the workbook actually renders in, which is this once anything has been authored over it.
- `setDefaultFont(font: Font): void;` — Author the workbook's default font — the face, size and colour every cell with no font of its own renders in, **empty cells included**. Merges into whatever the workbook already had, so `setDefaultFont({size: 14})` keeps the resolved face and changes only the size, and calling it twice accumulates. This is the one knob that reaches a cell no row or column default can: an untouched cell in an unformatted column. It writes the styles part's font 0 and **nothing else** — in particular it does not rewrite the theme's body typeface. The dependency runs the other way: with no default font authored, font 0 follows `themeFonts`'s minor face, so `setTheme({fonts: {minor}})` already reaches every unstyled cell and needs no second call here. See `defaultFont` for the full chain.
- `get defaultFont(): Font;` — The font every cell that names none of its own renders in, resolved and complete — what the writer emits as font id 0. Never `undefined`: a workbook always renders in *some* face, and the chain below always reaches one. ``` authored default font > authored theme body face > the source file's font 0 > theme body face ``` The two authored levels outrank the file because authoring is an explicit act; between them `setDefaultFont` wins on the face because it names font 0 outright while `setTheme` names it only by implication. With **nothing** authored the file's own font 0 passes through verbatim — deliberately, because a producer resolves that face by script and we do not: Excel writes `等线` as font 0 under a theme whose latin body face is `Calibri`, and re-deriving would silently rewrite it. `family` and `scheme` describe the *theme's* body face, so they are carried exactly while the resolved face still is that face and dropped when a caller names another — which is also what Excel writes: a font 0 naming a non-theme face carries no `<scheme>` at all. Either may be stated outright, in which case the caller's word stands.
- `authoredThemeXml(): string | undefined;` — The theme part text this workbook should write, or `undefined` when nothing was authored and the source theme (or the writer's default) should ride through untouched. Authoring generates *over* the existing part rather than from scratch — see `applyThemeOverrides` — so a preserved theme keeps its format scheme, its unauthored slots' exact encoding, and the relationships it carries.
- `resolveColor(color: Color): string | undefined;` — Resolve a colour reference to a concrete 8-hex ARGB string, or `undefined` when it does not resolve to a fixed colour — an `auto` colour, one of the two system indexed colours, or a theme slot this workbook's scheme does not declare. This is a *derived* view, not a rewrite: the `Color` stays exactly as its file encoded it, so a round-trip re-emits `theme="4" tint="0.4"` rather than a literal ARGB. Resolving into the model would sever every cell's link to the theme, so recolouring the workbook would stop working, and would inflate the styles table with one distinct colour per shade. A `theme` reference resolves through `themeColors`; an `indexed` one through the workbook's custom `<indexedColors>` palette when it declares one, else the built-in legacy palette. A `tint` is applied last.
- `get namedStyles(): readonly NamedCellStyle[];` — The named cell styles, in index order (index 0 is Normal); empty when only the default exists.
- `addPerson(person: Person): void;` — Register an identity a threaded comment can name — an author, or someone `@mentioned` in a message. A message reaches it by `Comment.personId`, a mention by `Mention.personId`. Keyed by `Person.id` alone, so registering the same id twice replaces the entry rather than adding a second: the id is the identity. Registering the same human twice under *different* ids is legitimate and is what Excel itself does — see `restorePersons`. The id is normalised to the brace-wrapped upper-case GUID form the format requires, so a `crypto.randomUUID()` is accepted as-is.
- `get persons(): readonly Person[];` — The registered threaded-comment identities, in the order they were read. That order carries no meaning — Excel re-sorts the registry by person id when it saves — so nothing may depend on it.
- `getPerson(id: string): Person | undefined;` — Look up a registered identity by its `Person.id`, or `undefined` if the registry has none.
- `addImage(options: AddImageOptions): number;` — Register a picture on the workbook and return its numeric id. Pass the id to `Worksheet.addImage` to anchor the picture to a sheet; the same id may be anchored on any number of sheets and positions, and the bytes are still stored only once.
- `get media(): readonly WorkbookImage[];` — The registered images, indexed by the id `addImage` returned.
- `getImage(id: number): WorkbookImage | undefined;` — Look up a registered image by its id, or `undefined` if no image carries that id.
- `get definedNames(): readonly DefinedName[];` — The workbook's defined names, in the order they were registered.
- `defineName(definedName: DefinedName): void;` — Register a defined name on the workbook.
- `addWorksheet(name: string, options: AddWorksheetOptions = {}): Worksheet;` — Create a worksheet and append it to the workbook.
- `getWorksheet(nameOrId: string | number): Worksheet | undefined;` — Look up a worksheet by name (case-insensitive) or by numeric id.

---

### `WorkbookProperties`

<sub>interface</sub>

Document-level metadata written to the package's core properties.

```ts
interface WorkbookProperties {
    creator?: string;
    lastModifiedBy?: string;
    created?: Date;
    modified?: Date;
}
```

---

### `WorkbookView`

<sub>interface</sub>

The workbook's saved window state — OOXML's `<workbookView>`, the single entry of `<bookViews>`.

This is the rect a consumer restores the document window to, and the layout every pane geometry is
computed against: a frozen split is positioned within it. `activeTab` names the sheet whose tab is
selected on open.

The position and size are in twips (1/20 of a point), Excel's window unit. A slightly negative
`x`/`y` is normal and is what Excel itself writes — a maximised window's frame sits just outside the
work area.

```ts
interface WorkbookView {
    x: number;
    y: number;
    width: number;
    height: number;
    activeTab: number;
    visibility?: 'visible' | 'hidden' | 'veryHidden';
    minimized?: boolean;
}
```
