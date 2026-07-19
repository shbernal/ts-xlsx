# Workbook

<!-- Generated from the public types by `npm run docs`. Do not edit by hand. -->

### `AddWorksheetOptions`

<sub>interface</sub>

```ts
interface AddWorksheetOptions {
    readonly state?: WorksheetState['state'];
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
  restoreDifferentialStyles(fragments: readonly string[]): void;
  get differentialStyles(): readonly string[];
  restoreIndexedColors(fragments: readonly string[]): void;
  get indexedColors(): readonly string[];
  restoreNamedStyles(styles: readonly NamedCellStyle[]): void;
  get namedStyles(): readonly NamedCellStyle[];
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
- `restoreDifferentialStyles(fragments: readonly string[]): void;` — Reinstate the differential-style (`<dxfs>`) table read from a file — the deserialization counterpart the writer re-emits verbatim. Each entry is one `<dxf>…</dxf>` fragment, preserved as opaque XML so a conditional-formatting rule's `dxfId` (an index into this table) stays valid on re-write. Replaces any table already held.
- `get differentialStyles(): readonly string[];` — The preserved differential-style (`<dxfs>`) fragments, in index order.
- `restoreIndexedColors(fragments: readonly string[]): void;` — Reinstate the custom indexed-color palette (`<colors><indexedColors>`) read from a file — each entry a verbatim `<rgbColor rgb="…"/>` fragment — so a colour referenced by `indexed="…"` keeps its intended RGB on re-write instead of the palette being dropped and the colour shifting to a default-palette entry. Replaces any palette already held.
- `get indexedColors(): readonly string[];` — The preserved custom indexed-color palette, in index order; empty when the default palette rules.
- `restoreNamedStyles(styles: readonly NamedCellStyle[]): void;` — Reinstate the named cell styles (`cellStyleXfs`/`cellStyles`) read from a file, index for index, so a cell's link to a named style (its `xfId`) stays valid on re-write. Index 0 is the Normal default. Replaces any table already held.
- `get namedStyles(): readonly NamedCellStyle[];` — The named cell styles, in index order (index 0 is Normal); empty when only the default exists.
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
