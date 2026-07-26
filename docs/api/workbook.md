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
  fullCalcOnLoad = false;
  protection: WorkbookProtection | undefined = undefined;
  get worksheets(): readonly Worksheet[];
  addPreservedReference(reference: PreservedWorkbookReference): void;
  get preservedReferences(): readonly PreservedWorkbookReference[];
  addPreservedRootReference(reference: PreservedRootReference): void;
  get preservedRootReferences(): readonly PreservedRootReference[];
  get customUI(): readonly CustomUiDocument[];
  get vbaProject(): VbaProject | undefined;
  get vbaProjectBytes(): Uint8Array | undefined;
  set vbaProjectBytes(bytes: Uint8Array | undefined);
  get vbaProjectSigned(): boolean;
  get vbaProjectSignatures(): readonly VbaProjectSignature[];
  removeVbaModule(name: string): void;
  addVbaReference(ref: VbaLibraryReference): void;
  restoreDifferentialStyles(fragments: readonly string[]): void;
  get differentialStyles(): readonly string[];
  restoreIndexedColors(fragments: readonly string[]): void;
  get indexedColors(): readonly string[];
  restoreNamedStyles(styles: readonly NamedCellStyle[]): void;
  get namedStyles(): readonly NamedCellStyle[];
  restorePersons(persons: readonly Person[]): void;
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

- `fullCalcOnLoad = false;` — Ask consuming spreadsheet apps to recalculate every formula when the file is opened, rather than trusting the cached results stored with each formula cell. Set this when the producer cannot compute formula results itself — the OOXML `fullCalcOnLoad` flag. Off by default, so a workbook whose cached results are authoritative stays unmarked.
- `protection: WorkbookProtection | undefined = undefined;` — Workbook-level structure/window protection — the OOXML `<workbookProtection>` element. Absent by default (an unprotected workbook). Set it to lock the workbook shell, or leave it as read from a file so a protected workbook stays locked across a passthrough save rather than being silently unlocked. Distinct from a worksheet's own `protect()`, which guards a single sheet's cells.
- `get worksheets(): readonly Worksheet[];` — The worksheets in insertion order.
- `addPreservedReference(reference: PreservedWorkbookReference): void;` — Record a workbook-level preserved reference (a pivot or slicer cache) read from a file.
- `get preservedReferences(): readonly PreservedWorkbookReference[];` — The workbook-level preserved references, in the order they were read.
- `addPreservedRootReference(reference: PreservedRootReference): void;` — Record a package-root preserved reference (a customUI ribbon part, custom props) read from a file.
- `get preservedRootReferences(): readonly PreservedRootReference[];` — The package-root preserved references, in the order they were read.
- `get customUI(): readonly CustomUiDocument[];` — The ribbon customisations decoded from this workbook's `customUI` parts — `customUI.xml` (Office 2007) and/or `customUI14.xml` (Office 2010+), in the order their root relationships were read. Each `CustomUiDocument` is tagged with its dialect and exposes the parsed `<ribbon>` tree. Empty for a workbook that customises no ribbon. This is a **read-only view** over parts the writer already round-trips verbatim — mutating the returned objects changes nothing on write; the original `customUI` XML is re-emitted byte-for-byte regardless. Parsed lazily on first access and memoised.
- `get vbaProject(): VbaProject | undefined;` — The VBA project decoded from this workbook's preserved `vbaProject.bin`, or `undefined` for a workbook with no macros. This is a **read-only view** over the bytes the writer already round-trips verbatim — mutating the returned object changes nothing on write; the original macro blob is re-emitted byte-for-byte regardless. Parsed lazily on first access and memoised.
- `get vbaProjectBytes(): Uint8Array | undefined;` — The raw `vbaProject.bin` bytes attached to this workbook — the exact macro blob the writer will embed — or `undefined` for a workbook with no macros. The getter returns a defensive copy, so mutating it changes nothing on write. Assigning bytes attaches (or replaces) the macro project: the written package becomes macro-enabled and re-embeds these bytes verbatim. The bytes must be a well-formed VBA container (a CFB holding a `dir` stream); a malformed blob is rejected with `VbaParseError` rather than written out to produce a package Excel would flag for repair. This is the attach-blob path: copy a project between workbooks with `dst.vbaProjectBytes = src.vbaProjectBytes`, or import a `.bin` produced by another tool. Assigning `undefined` removes the project, reverting the workbook to a plain (non-macro) package. Replacing or removing the project also drops any digital signature the previous blob carried — a signature over the old bytes cannot validate new ones — so the result never advertises a broken signature.
- `get vbaProjectSigned(): boolean;` — Whether this workbook's VBA project carries a digital signature — `true` if any signature part is attached, `false` for an unsigned project or a workbook with no macros. This reflects the **presence** of a signature blob, not its cryptographic validity: the library neither parses the PKCS#7/CMS structure nor validates the certificate chain or signer. A `true` here means "a signature is attached," never "this signature is valid." Replacing or editing the project drops its signatures (a signature over the old bytes cannot validate new ones), so this reads `false` again after `vbaProjectBytes`, `removeVbaModule`, or `addVbaReference` mutates the project. See `vbaProjectSignatures` for the raw bytes and which generation(s) are present.
- `get vbaProjectSignatures(): readonly VbaProjectSignature[];` — The digital signatures attached to this workbook's VBA project, in the order their relationships are wired off `vbaProject.bin` — up to three generations (legacy, agile, V3) can coexist over the same project bytes. Empty for an unsigned project or a workbook with no macros. Each entry's `bytes` are the raw signature blob passed through verbatim; the library does not parse or verify them (see `vbaProjectSigned` on presence-vs-validity). Hand a blob to an external verifier if you need cryptographic validation — that is deliberately out of this library's scope.
- `removeVbaModule(name: string): void;` — Remove a standard module from this workbook's existing macro project, in place — a structural splice that leaves every remaining module's compiled p-code untouched (see `removeVbaModule`). Replacing the project also drops a stale signature, as `vbaProjectBytes` does. Only `procedural` and `class` modules can be removed this way — see `removeVbaModule` for why. To author or edit module *source* (which needs real compiled p-code), use the offline `tools/vba-compiler`, then attach its output via `vbaProjectBytes`.
- `addVbaReference(ref: VbaLibraryReference): void;` — Add a registered (COM type-library) reference to this workbook's existing macro project, in place. Every existing module, reference, and host-info record rides through unchanged (see `addVbaReference`). Replacing the project also drops a stale signature, as `vbaProjectBytes` does.
- `restoreDifferentialStyles(fragments: readonly string[]): void;` — Reinstate the differential-style (`<dxfs>`) table read from a file — the deserialization counterpart the writer re-emits verbatim. Each entry is one `<dxf>…</dxf>` fragment, preserved as opaque XML so a conditional-formatting rule's `dxfId` (an index into this table) stays valid on re-write. Replaces any table already held.
- `get differentialStyles(): readonly string[];` — The preserved differential-style (`<dxfs>`) fragments, in index order.
- `restoreIndexedColors(fragments: readonly string[]): void;` — Reinstate the custom indexed-color palette (`<colors><indexedColors>`) read from a file — each entry a verbatim `<rgbColor rgb="…"/>` fragment — so a colour referenced by `indexed="…"` keeps its intended RGB on re-write instead of the palette being dropped and the colour shifting to a default-palette entry. Replaces any palette already held.
- `get indexedColors(): readonly string[];` — The preserved custom indexed-color palette, in index order; empty when the default palette rules.
- `restoreNamedStyles(styles: readonly NamedCellStyle[]): void;` — Reinstate the named cell styles (`cellStyleXfs`/`cellStyles`) read from a file, index for index, so a cell's link to a named style (its `xfId`) stays valid on re-write. Index 0 is the Normal default. Replaces any table already held.
- `get namedStyles(): readonly NamedCellStyle[];` — The named cell styles, in index order (index 0 is Normal); empty when only the default exists.
- `restorePersons(persons: readonly Person[]): void;` — Reinstate the threaded-comment identity registry (`xl/persons/person.xml`) read from a file — the authors and mentioned people a comment thread's messages point at. Replaces any registry already held. Entries are keyed by `Person.id` and by nothing else. A single human legitimately owns several entries: Excel interns a *mentioned* identity as its own `providerId="PeoplePicker"` entry beside that person's `providerId="AD"` authoring entry — same `displayName`, same `userId`, a different id — and points the mention at the new one. Collapsing entries by name or `userId` would merge those two and silently re-point every mention at the wrong identity.
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
