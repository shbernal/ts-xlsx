// The workbook: the top of the model and the entry point of the public API.
//
// It owns its worksheets and the document-level properties. Sheet identity follows
// Excel's rules — names are unique case-insensitively, bounded in length, and free
// of the characters Excel forbids — so an invalid book cannot be constructed in the
// first place, rather than failing only at write time.

import {type CustomUiDocument, isCustomUiRelType, parseCustomUi} from '../customui/index.ts';
import {AuthoringError} from '../errors.ts';
import {
  addVbaReference,
  parseVbaProject,
  removeVbaModule,
  VBA_PROJECT_CONTENT_TYPE,
  VBA_PROJECT_PART_PATH,
  VBA_PROJECT_REL_TYPE,
  VbaAuthorError,
  type VbaLibraryReference,
  type VbaProject,
  type VbaProjectSignature,
  vbaProjectSignatureKind,
} from '../vba/index.ts';
import {resolveColor} from './color-resolution.ts';
import {commentThreadGuid, type Person} from './comment-thread.ts';
import {replaceContents} from './containers.ts';
import {normalizeImageExtension, type WorkbookImage} from './image.ts';
import {INTERNAL} from './internal.ts';
import type {PreservedPart, PreservedRootReference} from './preserved.ts';
import type {Color, NamedCellStyle, TableStyleTable} from './style.ts';
import {checkTableStyle, type TableStyle} from './table-style.ts';
import {
  applyThemeOverrides,
  DEFAULT_THEME_COLOR_SCHEME,
  DEFAULT_THEME_FONTS,
  DEFAULT_THEME_XML,
  parseThemeColorScheme,
  parseThemeFontScheme,
  THEME_COLOR_SLOTS,
  type ThemeColorScheme,
  type ThemeColorSlot,
  type ThemeFontScheme,
  type ThemeOverrides,
} from './theme.ts';
import type {WorkbookProtection} from './workbook-protection.ts';
import {Worksheet, type WorksheetState} from './worksheet.ts';

/**
 * A workbook-level reference to package content the model does not model — a pivot cache
 * (`pivotCacheDefinition`) or a slicer cache (`slicerCache`) — preserved verbatim across a round-trip
 * instead of being dropped. `relType` is the workbook relationship Type URI to re-emit; `entryPath`
 * is the part it points at; `parts` is the transitive closure that reference reaches (the entry
 * included). `pivotCacheId` carries the `<pivotCache cacheId>` a pivot cache is registered under in
 * the workbook's `<pivotCaches>`, so the wiring a pivot table resolves its cache through is re-emitted
 * too; it is absent for a slicer cache, which the workbook lists in an extension block instead.
 * `externalReferenceIndex` is the 0-based position of an `externalLink` within the workbook's
 * `<externalReferences>` — the `[n]` a formula or defined name resolves an external cell through — so
 * the block is re-emitted in the original order and every `[n]` still points at the same linked
 * workbook; it is absent for a pivot/slicer cache.
 */
export interface PreservedWorkbookReference {
  readonly relType: string;
  readonly entryPath: string;
  readonly parts: readonly PreservedPart[];
  readonly pivotCacheId?: string;
  readonly externalReferenceIndex?: number;
}

/**
 * The workbook's theme part, captured verbatim from a source package — the `<clrScheme>`,
 * `<fontScheme>`, and `<fmtScheme>` every `theme="n"` colour reference and every `scheme="major|minor"`
 * font in the file resolves against.
 *
 * Held opaquely, like {@link Workbook.restoreDifferentialStyles}'s `<dxf>` fragments: the model does
 * not interpret the theme, it only refuses to destroy it. `entryPath` is where the source package held
 * the part (reached through the workbook's `.../theme` relationship, so not necessarily
 * `xl/theme/theme1.xml`), and `parts` is the transitive closure it reaches — the entry included. The
 * closure matters because a theme can carry its own relationships: a picture used as a themed fill is
 * wired by an `r:embed` into the theme's rels part, and re-emitting the theme without it would leave
 * that reference dangling.
 */
export interface PreservedTheme {
  readonly entryPath: string;
  readonly parts: readonly PreservedPart[];
}

/**
 * The workbook's saved window state — OOXML's `<workbookView>`, the single entry of `<bookViews>`.
 *
 * This is the rect a consumer restores the document window to, and the layout every pane geometry is
 * computed against: a frozen split is positioned within it. `activeTab` names the sheet whose tab is
 * selected on open.
 *
 * The position and size are in twips (1/20 of a point), Excel's window unit. A slightly negative
 * `x`/`y` is normal and is what Excel itself writes — a maximised window's frame sits just outside the
 * work area.
 */
export interface WorkbookView {
  /** Left edge of the document window, in twips. */
  x: number;
  /** Top edge of the document window, in twips. */
  y: number;
  /** Window width, in twips. */
  width: number;
  /** Window height, in twips. */
  height: number;
  /** 0-based index into {@link Workbook.worksheets} of the sheet selected on open. */
  activeTab: number;
  /** Window visibility; omit for a normally visible window. */
  visibility?: 'visible' | 'hidden' | 'veryHidden';
  /** Whether the document window opens minimised; omit for a restored window. */
  minimized?: boolean;
}

/**
 * The window geometry a workbook starts from — the values desktop Excel writes for its own default
 * window.
 *
 * A default is emitted rather than the element left out because Excel writes `<bookViews>` into every
 * file it saves and consumers lay panes out against that rect. With no view at all the frozen-pane
 * split is computed against an uninitialised window, and the frozen region can stay unpainted until
 * some later event forces a relayout.
 */
export const DEFAULT_WORKBOOK_VIEW = {
  x: -110,
  y: -110,
  width: 19420,
  height: 12220,
  activeTab: 0,
} as const satisfies WorkbookView;

/** Document-level metadata written to the package's core properties. */
export interface WorkbookProperties {
  creator?: string;
  lastModifiedBy?: string;
  created?: Date;
  modified?: Date;
}

/**
 * A named reference in the workbook — the entries Excel surfaces in its Name Manager. A name maps
 * an identifier to a formula (`refersTo`), most often a cell range like `Sheet1!$A$1:$B$2` but
 * possibly any formula. A name is global to the workbook unless it names a sheet in {@link scope},
 * which restricts it to that sheet and lets another sheet reuse the same name independently.
 */
export interface DefinedName {
  /** The name as typed in a formula, e.g. `"TaxRate"`. Built-in names carry an `_xlnm.` prefix. */
  readonly name: string;
  /** The formula the name resolves to, e.g. `"Sheet1!$A$1:$B$2"`. */
  readonly refersTo: string;
  /** The sheet the name is scoped to; omit for a workbook-global name. */
  readonly scope?: string;
  /** A human note shown beside the name in Excel's Name Manager. */
  readonly comment?: string;
  /** Hide the name from the Name Manager UI without removing it. */
  readonly hidden?: boolean;
}

export interface AddWorksheetOptions {
  readonly state?: WorksheetState['state'];
}

/** A picture registered on the workbook, ready to be anchored to a worksheet. */
export interface AddImageOptions {
  /** The image bytes. */
  readonly buffer: Uint8Array;
  /** The file kind — `"png"`, `"jpeg"`/`"jpg"`, `"gif"`, … A leading dot or a URL query string is
   * tolerated and stripped; omit it entirely to infer the kind from the bytes' magic number. */
  readonly extension?: string;
}

const MAX_SHEET_NAME_LENGTH = 31;
// Excel rejects these in a sheet name, plus a leading/trailing apostrophe.
const INVALID_SHEET_NAME_CHARS = /[*?:\\/[\]]/;

export class Workbook {
  readonly properties: WorkbookProperties = {};

  /**
   * The workbook's window state — position, size, and the selected sheet. Always present (see
   * {@link DEFAULT_WORKBOOK_VIEW} for why it is defaulted rather than left unset) and always written.
   * Reading a file replaces it with that file's saved geometry, so a round-trip restores the window
   * the author left rather than stamping ours over it.
   */
  readonly view: WorkbookView = {...DEFAULT_WORKBOOK_VIEW};

  /**
   * Ask consuming spreadsheet apps to recalculate every formula when the file is opened, rather than
   * trusting the cached results stored with each formula cell. Set this when the producer cannot
   * compute formula results itself — the OOXML `fullCalcOnLoad` flag. Off by default, so a workbook
   * whose cached results are authoritative stays unmarked.
   */
  fullCalcOnLoad = false;

  /**
   * Workbook-level structure/window protection — the OOXML `<workbookProtection>` element. Absent by
   * default (an unprotected workbook). Set it to lock the workbook shell, or leave it as read from a
   * file so a protected workbook stays locked across a passthrough save rather than being silently
   * unlocked. Distinct from a worksheet's own `protect()`, which guards a single sheet's cells.
   */
  protection: WorkbookProtection | undefined = undefined;

  readonly #worksheets: Worksheet[] = [];
  #nextSheetId = 1;

  // Media is shared workbook-wide: a worksheet anchors an image by its registry index, so one
  // picture used on several sheets is stored once.
  readonly #media: WorkbookImage[] = [];

  readonly #definedNames: DefinedName[] = [];

  // Differential styles (`<dxfs>`) are a workbook-level table in styles.xml that conditional
  // formatting references by index. The library models the classic scale rules directly but preserves
  // the dxf table as opaque XML fragments, so a rule that references a dxfId (a highlight fill, a
  // custom number format) keeps a valid target across a read/write cycle instead of dangling.
  readonly #dxfs: string[] = [];

  // Named cell styles (`cellStyleXfs`/`cellStyles` in styles.xml) — the shared, named formatting layer
  // a cell links to by index. Preserved so a cell whose fill/font/… lives only in a named style keeps
  // that style, and the link, across a round-trip. Empty when a file declares nothing beyond the
  // default Normal style, in which case the writer emits just that default.
  readonly #namedStyles: NamedCellStyle[] = [];

  // A custom indexed-color palette (`<colors><indexedColors>` in styles.xml) read from a file, each
  // entry a verbatim `<rgbColor rgb="…"/>` fragment. Preserved so an `indexed="…"` colour reference
  // keeps its intended RGB across a round-trip instead of resolving to a different default-palette
  // entry. Empty for a workbook that never overrode the palette.
  readonly #indexedColors: string[] = [];

  // The most-recently-used colour swatches (`<colors><mruColors>` in styles.xml), each a verbatim
  // `<color rgb="…"/>` fragment. The author's own working set of colours; dropping it on a re-write
  // quietly resets a habit. Empty for a workbook that never picked a custom colour.
  readonly #mruColors: string[] = [];

  // The custom table-style definitions (`<tableStyles>` in styles.xml), each `<tableStyle>` kept
  // verbatim, plus the gallery names the file nominates as the default for a new table and pivot. A
  // table's `tableStyleInfo/@name` can name one of these definitions, so dropping the block leaves
  // that reference dangling and the table renders unstyled.
  #tableStyles: TableStyleTable = {styles: []};

  // The theme part read from a file, kept verbatim with the closure of parts it reaches. The writer
  // emits its own default theme for a workbook that has none, so without this a branded theme would be
  // overwritten by that default and every `theme="n"` colour in the file would silently re-render.
  // Undefined for a workbook authored from scratch, or read from a package declaring no theme.
  #theme: PreservedTheme | undefined;

  // Workbook-level references to package content the model does not interpret (pivot caches, slicer
  // caches), captured verbatim on read so a round-trip re-emits them rather than dropping the pivots
  // and slicers they back. Empty for a workbook authored from scratch.
  readonly #preservedReferences: PreservedWorkbookReference[] = [];

  // The threaded-comment identity registry (`xl/persons/person.xml`), the workbook-level table every
  // message resolves its author through and every @mention its target. Keyed by person id — see
  // `restorePersons` for why nothing else will do. Empty for a workbook with no threaded comments.
  readonly #persons = new Map<string, Person>();

  /** The worksheets in insertion order. */
  get worksheets(): readonly Worksheet[] {
    return this.#worksheets;
  }

  /**
   * The 0-based index of the active sheet: {@link WorkbookView.activeTab} resolved against the sheets
   * that actually exist. Exactly one sheet is always active — an out-of-range tab (a caller's stale
   * index, or a file whose sheet was removed after the view was saved) falls back to the first sheet
   * rather than to none, because a package where no sheet is selected gives the consumer no view to
   * initialise on open.
   */
  get activeTabIndex(): number {
    const {activeTab} = this.view;
    return Number.isInteger(activeTab) && activeTab > 0 && activeTab < this.#worksheets.length
      ? activeTab
      : 0;
  }

  /** The workbook-level preserved references, in the order they were read. */
  get preservedReferences(): readonly PreservedWorkbookReference[] {
    return this.#preservedReferences;
  }

  // Package-root references to unmodeled content wired from `_rels/.rels` (the ribbon customUI parts,
  // custom document properties, a thumbnail), captured verbatim on read so a round-trip re-declares
  // them in the regenerated root rels rather than dropping them. Empty for a workbook authored from
  // scratch.
  readonly #preservedRootReferences: PreservedRootReference[] = [];

  /** The package-root preserved references, in the order they were read. */
  get preservedRootReferences(): readonly PreservedRootReference[] {
    return this.#preservedRootReferences;
  }

  // Lazily-parsed customUI ribbon view. `#customUiParsed` distinguishes "not yet parsed" from a genuine
  // "no ribbon parts" (an empty array), so a ribbon-free workbook is not re-scanned on every access.
  #customUiParsed = false;
  #customUI: readonly CustomUiDocument[] = [];

  /**
   * The ribbon customisations decoded from this workbook's `customUI` parts — `customUI.xml` (Office
   * 2007) and/or `customUI14.xml` (Office 2010+), in the order their root relationships were read. Each
   * {@link CustomUiDocument} is tagged with its dialect and exposes the parsed `<ribbon>` tree. Empty
   * for a workbook that customises no ribbon.
   *
   * This is a **read-only view** over parts the writer already round-trips verbatim — mutating the
   * returned objects changes nothing on write; the original `customUI` XML is re-emitted byte-for-byte
   * regardless. Parsed lazily on first access and memoised.
   *
   * @throws {@link CustomUiParseError} if a `customUI` part is present but its XML is malformed.
   */
  get customUI(): readonly CustomUiDocument[] {
    if (!this.#customUiParsed) {
      this.#customUI = this.#preservedRootReferences
        .filter((ref) => isCustomUiRelType(ref.relType))
        .flatMap((ref) => {
          const bytes = ref.parts.find((part) => part.path === ref.entryPath)?.bytes;
          return bytes === undefined ? [] : [parseCustomUi(bytes)];
        });
      this.#customUiParsed = true;
    }
    return this.#customUI;
  }

  // Lazily-decoded macro source. `#vbaParsed` distinguishes "not yet decoded" from a genuine "no
  // macros" (`undefined`) result, so a macro-free workbook is not re-probed on every access.
  #vbaParsed = false;
  #vbaProject: VbaProject | undefined = undefined;

  /**
   * The VBA project decoded from this workbook's preserved `vbaProject.bin`, or `undefined` for a
   * workbook with no macros. This is a **read-only view** over the bytes the writer already round-trips
   * verbatim — mutating the returned object changes nothing on write; the original macro blob is
   * re-emitted byte-for-byte regardless. Parsed lazily on first access and memoised.
   *
   * @throws {@link VbaParseError} if a macro project is present but its `vbaProject.bin` is malformed.
   */
  get vbaProject(): VbaProject | undefined {
    if (!this.#vbaParsed) {
      const bytes = this.#vbaProjectEntry()?.bytes;
      this.#vbaProject = bytes ? parseVbaProject(bytes) : undefined;
      this.#vbaParsed = true;
    }
    return this.#vbaProject;
  }

  /**
   * The raw `vbaProject.bin` bytes attached to this workbook — the exact macro blob the writer will
   * embed — or `undefined` for a workbook with no macros. The getter returns a defensive copy, so
   * mutating it changes nothing on write.
   *
   * Assigning bytes attaches (or replaces) the macro project: the written package becomes
   * macro-enabled and re-embeds these bytes verbatim. The bytes must be a well-formed VBA container
   * (a CFB holding a `dir` stream); a malformed blob is rejected with {@link VbaParseError} rather
   * than written out to produce a package Excel would flag for repair. This is the attach-blob path:
   * copy a project between workbooks with `dst.vbaProjectBytes = src.vbaProjectBytes`, or import a
   * `.bin` produced by another tool. Assigning `undefined` removes the project, reverting the workbook
   * to a plain (non-macro) package.
   *
   * Replacing or removing the project also drops any digital signature the previous blob carried — a
   * signature over the old bytes cannot validate new ones — so the result never advertises a broken
   * signature.
   */
  get vbaProjectBytes(): Uint8Array | undefined {
    return this.#vbaProjectEntry()?.bytes.slice();
  }

  set vbaProjectBytes(bytes: Uint8Array | undefined) {
    // Validate before touching any state: a malformed blob must fail closed and leave the existing
    // project intact, never half-remove it. Only past this point do we mutate.
    if (bytes !== undefined) parseVbaProject(bytes);

    // Drop any existing project; its whole closure goes, taking a now-stale signature part with it. A
    // fresh reference then mirrors exactly what the reader captures for a macro workbook, so the writer
    // emits a byte-identical macro-enabled package with no writer changes.
    replaceContents(
      this.#preservedReferences,
      this.#preservedReferences.filter((r) => !r.relType.endsWith('/vbaProject')),
    );
    if (bytes !== undefined) {
      this.#preservedReferences.push({
        relType: VBA_PROJECT_REL_TYPE,
        entryPath: VBA_PROJECT_PART_PATH,
        parts: [
          {
            path: VBA_PROJECT_PART_PATH,
            contentType: VBA_PROJECT_CONTENT_TYPE,
            bytes: bytes.slice(),
            rels: [],
          },
        ],
      });
    }
    this.#vbaParsed = false;
    this.#vbaProject = undefined;
  }

  /**
   * Whether this workbook's VBA project carries a digital signature — `true` if any signature part is
   * attached, `false` for an unsigned project or a workbook with no macros.
   *
   * This reflects the **presence** of a signature blob, not its cryptographic validity: the library
   * neither parses the PKCS#7/CMS structure nor validates the certificate chain or signer. A `true`
   * here means "a signature is attached," never "this signature is valid." Replacing or editing the
   * project drops its signatures (a signature over the old bytes cannot validate new ones), so this
   * reads `false` again after {@link vbaProjectBytes}, {@link removeVbaModule}, or
   * {@link addVbaReference} mutates the project. See {@link vbaProjectSignatures} for the raw bytes and
   * which generation(s) are present.
   */
  get vbaProjectSigned(): boolean {
    return this.#vbaSignatures().length > 0;
  }

  /**
   * The digital signatures attached to this workbook's VBA project, in the order their relationships
   * are wired off `vbaProject.bin` — up to three generations (legacy, agile, V3) can coexist over the
   * same project bytes. Empty for an unsigned project or a workbook with no macros.
   *
   * Each entry's `bytes` are the raw signature blob passed through verbatim; the library does not parse
   * or verify them (see {@link vbaProjectSigned} on presence-vs-validity). Hand a blob to an external
   * verifier if you need cryptographic validation — that is deliberately out of this library's scope.
   */
  get vbaProjectSignatures(): readonly VbaProjectSignature[] {
    return this.#vbaSignatures();
  }

  // Walk the VBA project's preserved closure for its signature parts — each reached by a signature
  // relationship off `vbaProject.bin`. Computed on each access rather than memoised: the closure is
  // small and already in memory, and recomputing sidesteps a cache that a signature-dropping mutation
  // (`vbaProjectBytes` replace, module remove, reference add) would otherwise have to invalidate.
  #vbaSignatures(): readonly VbaProjectSignature[] {
    const ref = this.#vbaProjectRef();
    const entry = ref?.parts.find((p) => p.path === ref.entryPath);
    if (ref === undefined || entry === undefined) return [];
    const partByPath = new Map(ref.parts.map((p) => [p.path, p]));
    const signatures: VbaProjectSignature[] = [];
    for (const rel of entry.rels) {
      const kind = vbaProjectSignatureKind(rel.type);
      const part = kind === undefined ? undefined : partByPath.get(rel.targetPath);
      if (kind !== undefined && part !== undefined) {
        signatures.push({kind, bytes: part.bytes.slice()});
      }
    }
    return signatures;
  }

  /**
   * Remove a standard module from this workbook's existing macro project, in place — a structural splice
   * that leaves every remaining module's compiled p-code untouched (see {@link removeVbaModule}).
   * Replacing the project also drops a stale signature, as {@link vbaProjectBytes} does.
   *
   * Only `procedural` and `class` modules can be removed this way — see {@link removeVbaModule} for why.
   * To author or edit module *source* (which needs real compiled p-code), use the offline
   * `tools/vba-compiler`, then attach its output via {@link vbaProjectBytes}.
   *
   * @throws {@link VbaAuthorError} if the workbook has no macro project, or `name` is not in the project,
   *   or names a `document`/`designer` module.
   * @throws {@link VbaParseError} if the attached `vbaProject.bin` is malformed.
   */
  removeVbaModule(name: string): void {
    const bytes = this.vbaProjectBytes;
    if (bytes === undefined) {
      throw new VbaAuthorError('workbook has no VBA project to remove a module from');
    }
    this.vbaProjectBytes = removeVbaModule(bytes, name);
  }

  /**
   * Add a registered (COM type-library) reference to this workbook's existing macro project, in place.
   * Every existing module, reference, and host-info record rides through unchanged (see
   * {@link addVbaReference}). Replacing the project also drops a stale signature, as
   * {@link vbaProjectBytes} does.
   *
   * @throws {@link VbaAuthorError} if the workbook has no macro project, or any field of `ref` is invalid
   *   (see {@link VbaLibraryReference}).
   * @throws {@link VbaParseError} if the attached `vbaProject.bin` is malformed.
   */
  addVbaReference(ref: VbaLibraryReference): void {
    const bytes = this.vbaProjectBytes;
    if (bytes === undefined) {
      throw new VbaAuthorError('workbook has no VBA project to add a reference to');
    }
    this.vbaProjectBytes = addVbaReference(bytes, ref);
  }

  #vbaProjectRef(): PreservedWorkbookReference | undefined {
    return this.#preservedReferences.find((r) => r.relType.endsWith('/vbaProject'));
  }

  #vbaProjectEntry(): PreservedPart | undefined {
    const ref = this.#vbaProjectRef();
    return ref?.parts.find((p) => p.path === ref.entryPath);
  }

  /** The preserved differential-style (`<dxfs>`) fragments, in index order. */
  get differentialStyles(): readonly string[] {
    return this.#dxfs;
  }

  /** The preserved custom indexed-color palette, in index order; empty when the default palette rules. */
  get indexedColors(): readonly string[] {
    return this.#indexedColors;
  }

  /** The preserved most-recently-used colour swatches, in order; empty when the file declared none. */
  get mruColors(): readonly string[] {
    return this.#mruColors;
  }

  /** The preserved `<tableStyles>` block; `styles` is empty when the file declared no custom style. */
  get tableStyles(): TableStyleTable {
    return this.#tableStyles;
  }

  // Table styles authored on this workbook, keyed by name so registering the same name twice replaces
  // rather than duplicates — two definitions sharing a name leave a table's reference ambiguous.
  readonly #customTableStyles = new Map<string, TableStyle>();

  /**
   * Register a custom table style — a named look a table applies to itself by putting that name in
   * {@link TableStyleInfo.name}, exactly as it would name one of Excel's built-in gallery styles.
   *
   * ```ts
   * workbook.addTableStyle({
   *   name: 'Harbour',
   *   elements: {
   *     wholeTable: {border: {top: {style: 'thin'}, bottom: {style: 'thin'}}},
   *     headerRow: {font: {bold: true, color: {argb: 'FFFFFFFF'}},
   *                 fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFBB2649'}}},
   *     firstRowStripe: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFF6E7EB'}}},
   *   },
   * });
   * sheet.addTable({name: 'Cargo', ref: 'A1:B3', columns, style: {name: 'Harbour'}});
   * ```
   *
   * Each element's formatting is interned into the workbook's shared differential-style table, so two
   * elements — or a conditional-formatting rule — that paint the same way share one entry.
   *
   * Registering a name a source file already defined **overrides** that definition rather than adding
   * a second one beside it.
   *
   * @throws {@link AuthoringError} if the name is empty, or an element carries a `size` outside the four stripe
   *   types, or a `size` is not a positive integer — see {@link checkTableStyle} for why those are
   *   refused here rather than silently dropped.
   */
  addTableStyle(style: TableStyle): void {
    checkTableStyle(style);
    this.#customTableStyles.set(style.name, style);
  }

  /** The table styles authored on this workbook, in registration order. */
  get customTableStyles(): readonly TableStyle[] {
    return [...this.#customTableStyles.values()];
  }

  /** The preserved theme part, or undefined when the workbook rides the library's default theme. */
  get themePart(): PreservedTheme | undefined {
    return this.#theme;
  }

  // The theme's colour scheme, decoded from the preserved part (and merged with any authored
  // overrides) on first use. Cached because resolving a colour is a per-cell operation and the part is
  // otherwise held as bytes; invalidated whenever the theme is replaced or authored.
  #themeColors: ThemeColorScheme | undefined;

  // Colour slots and typefaces the caller authored, merged over whatever the workbook already had.
  #authoredTheme: {colors: {-readonly [K in ThemeColorSlot]?: string}; fonts: ThemeFontScheme} = {
    colors: {},
    fonts: {},
  };

  /**
   * Author the workbook's theme: any subset of the twelve colour-scheme slots, and either of the two
   * typefaces. Merges into what the workbook already has, so branding one accent leaves the other
   * eleven slots alone, and calling it twice accumulates.
   *
   * This is the workbook-wide palette. A cell that names a colour as `theme="4"` — which is what Excel
   * writes whenever a user picks from the theme row of the colour picker — follows `accent1` here, so
   * one call restyles every such cell, chart and table style at once. Colours are `RRGGBB`; a leading
   * `#` and an 8-hex ARGB are both accepted and reduced, and anything else throws rather than writing
   * a value Excel silently renders as flat black.
   *
   * What it does **not** touch: the theme's format scheme — the gradient, line and effect styles that
   * give a theme its texture. Those ride through from the source theme (or the library's default)
   * untouched, because nobody hand-authors gradient stops from a spreadsheet API and regenerating them
   * would replace a designer's work with the Office default. For the same reason a slot left
   * unauthored keeps the source's own encoding, including the `<a:sysClr>` form Excel uses for
   * `dk1`/`lt1` so they follow the viewer's window colours.
   *
   * @throws {@link AuthoringError} if a colour is not 6 or 8 hexadecimal digits.
   */
  setTheme(overrides: ThemeOverrides): void {
    // Validated eagerly, by running the generation the writer will later run: a colour rejected at
    // write time would surface far from the call that supplied it.
    applyThemeOverrides(this.#baseThemeXml(), overrides);
    Object.assign(this.#authoredTheme.colors, overrides.colors ?? {});
    this.#authoredTheme.fonts = {...this.#authoredTheme.fonts, ...(overrides.fonts ?? {})};
    this.#themeColors = undefined;
  }

  /**
   * The colour scheme every `theme="n"` reference in this workbook resolves against — anything
   * {@link setTheme} authored, over the preserved theme's `<a:clrScheme>`, over the Office default.
   *
   * Note the slot *order*: `theme="0"` is `lt1` and `theme="1"` is `dk1`, which is not the order the
   * slots appear in the theme part. See {@link THEME_COLOR_SLOTS}.
   */
  get themeColors(): ThemeColorScheme {
    if (this.#themeColors === undefined) {
      const xml = this.#themeXml();
      // A theme that declares no scheme (or none this reader decodes) falls back to the Office
      // default rather than resolving nothing: the file still renders against *some* scheme, and the
      // default is the one the writer would have shipped.
      const parsed = xml === undefined ? {} : parseThemeColorScheme(xml);
      const base = Object.keys(parsed).length === 0 ? DEFAULT_THEME_COLOR_SCHEME : parsed;
      this.#themeColors = {...base, ...this.#authoredTheme.colors};
    }
    return this.#themeColors;
  }

  /** The theme's major (heading) and minor (body) typefaces, authored values over the source's. */
  get themeFonts(): ThemeFontScheme {
    const xml = this.#themeXml();
    const parsed = xml === undefined ? {} : parseThemeFontScheme(xml);
    const base = Object.keys(parsed).length === 0 ? DEFAULT_THEME_FONTS : parsed;
    return {...base, ...this.#authoredTheme.fonts};
  }

  /**
   * The theme part text this workbook should write, or `undefined` when nothing was authored and the
   * source theme (or the writer's default) should ride through untouched.
   *
   * Authoring generates *over* the existing part rather than from scratch — see
   * {@link applyThemeOverrides} — so a preserved theme keeps its format scheme, its unauthored slots'
   * exact encoding, and the relationships it carries.
   */
  authoredThemeXml(): string | undefined {
    const {colors, fonts} = this.#authoredTheme;
    if (Object.keys(colors).length === 0 && Object.keys(fonts).length === 0) return undefined;
    return applyThemeOverrides(this.#baseThemeXml(), {colors, fonts});
  }

  // The part authored overrides are applied on top of: the preserved source theme, else the default
  // one the writer would otherwise have emitted.
  #baseThemeXml(): string {
    return this.#themeXml() ?? DEFAULT_THEME_XML;
  }

  // The preserved theme part's text, decoded from the entry part of its closure.
  #themeXml(): string | undefined {
    const theme = this.#theme;
    if (theme === undefined) return undefined;
    const entry = theme.parts.find((part) => part.path === theme.entryPath);
    return entry === undefined ? undefined : new TextDecoder().decode(entry.bytes);
  }

  /**
   * Resolve a colour reference to a concrete 8-hex ARGB string, or `undefined` when it does not
   * resolve to a fixed colour — an `auto` colour, one of the two system indexed colours, or a theme
   * slot this workbook's scheme does not declare.
   *
   * This is a *derived* view, not a rewrite: the {@link Color} stays exactly as its file encoded it,
   * so a round-trip re-emits `theme="4" tint="0.4"` rather than a literal ARGB. Resolving into the
   * model would sever every cell's link to the theme, so recolouring the workbook would stop working,
   * and would inflate the styles table with one distinct colour per shade.
   *
   * A `theme` reference resolves through {@link themeColors}; an `indexed` one through the workbook's
   * custom `<indexedColors>` palette when it declares one, else the built-in legacy palette. A `tint`
   * is applied last.
   */
  resolveColor(color: Color): string | undefined {
    return resolveColor(color, {theme: this.themeColors, indexed: this.#indexedPalette()});
  }

  // The workbook's custom palette as plain ARGB strings. `#indexedColors` holds verbatim
  // `<rgbColor rgb="…"/>` fragments — the form the writer re-emits — so the value is read out here
  // rather than stored twice in two shapes that could drift.
  #indexedPalette(): readonly string[] {
    return this.#indexedColors.map((fragment) => /\brgb="([^"]*)"/.exec(fragment)?.[1] ?? '');
  }

  /** The named cell styles, in index order (index 0 is Normal); empty when only the default exists. */
  get namedStyles(): readonly NamedCellStyle[] {
    return this.#namedStyles;
  }

  /**
   * Register an identity a threaded comment can name — an author, or someone `@mentioned` in a message.
   * A message reaches it by {@link Comment.personId}, a mention by {@link Mention.personId}.
   *
   * Keyed by {@link Person.id} alone, so registering the same id twice replaces the entry rather than
   * adding a second: the id is the identity. Registering the same human twice under *different* ids is
   * legitimate and is what Excel itself does — see {@link restorePersons}. The id is normalised to the
   * brace-wrapped upper-case GUID form the format requires, so a `crypto.randomUUID()` is accepted as-is.
   *
   * @throws {SyntaxError} if the id is not a GUID.
   */
  addPerson(person: Person): void {
    const id = commentThreadGuid(person.id, 'a person id');
    this.#persons.set(id, {...person, id});
  }

  /**
   * The registered threaded-comment identities, in the order they were read. That order carries no
   * meaning — Excel re-sorts the registry by person id when it saves — so nothing may depend on it.
   */
  get persons(): readonly Person[] {
    return [...this.#persons.values()];
  }

  /** Look up a registered identity by its {@link Person.id}, or `undefined` if the registry has none. */
  getPerson(id: string): Person | undefined {
    return this.#persons.get(id);
  }

  /**
   * Register a picture on the workbook and return its numeric id. Pass the id to
   * {@link Worksheet.addImage} to anchor the picture to a sheet; the same id may be anchored on any
   * number of sheets and positions, and the bytes are still stored only once.
   */
  addImage(options: AddImageOptions): number {
    this.#media.push({
      extension: normalizeImageExtension(options.extension, options.buffer),
      data: options.buffer,
    });
    return this.#media.length - 1;
  }

  /** The registered images, indexed by the id {@link addImage} returned. */
  get media(): readonly WorkbookImage[] {
    return this.#media;
  }

  /** Look up a registered image by its id, or `undefined` if no image carries that id. */
  getImage(id: number): WorkbookImage | undefined {
    return this.#media[id];
  }

  /** The workbook's defined names, in the order they were registered. */
  get definedNames(): readonly DefinedName[] {
    return this.#definedNames;
  }

  /**
   * Register a defined name on the workbook.
   *
   * @throws {@link AuthoringError} if the name is empty, or if a {@link DefinedName.scope} is given that names no
   *   existing worksheet — a scoped name must target a sheet that is already part of the workbook.
   */
  defineName(definedName: DefinedName): void {
    if (definedName.name.length === 0) {
      throw new AuthoringError('a defined name cannot be empty');
    }
    if (definedName.scope !== undefined && this.getWorksheet(definedName.scope) === undefined) {
      throw new AuthoringError(
        `defined name "${definedName.name}" is scoped to unknown worksheet "${definedName.scope}"`,
      );
    }
    this.#definedNames.push(definedName);
  }

  /**
   * Create a worksheet and append it to the workbook.
   *
   * @throws {@link AuthoringError} if the name is empty, too long, contains a forbidden character,
   *   or collides (case-insensitively) with an existing sheet.
   */
  addWorksheet(name: string, options: AddWorksheetOptions = {}): Worksheet {
    this.#assertValidSheetName(name);
    const sheet = new Worksheet(name, this.#nextSheetId++, options.state ?? 'visible');
    this.#worksheets.push(sheet);
    return sheet;
  }

  /** Look up a worksheet by name (case-insensitive) or by numeric id. */
  getWorksheet(nameOrId: string | number): Worksheet | undefined {
    if (typeof nameOrId === 'number') {
      return this.#worksheets.find((sheet) => sheet.id === nameOrId);
    }
    const target = nameOrId.toLowerCase();
    return this.#worksheets.find((sheet) => sheet.name.toLowerCase() === target);
  }

  #assertValidSheetName(name: string): void {
    if (name.length === 0) {
      throw new AuthoringError('worksheet name cannot be empty');
    }
    if (name.length > MAX_SHEET_NAME_LENGTH) {
      throw new AuthoringError(
        `worksheet name "${name}" exceeds the ${MAX_SHEET_NAME_LENGTH}-character limit`,
      );
    }
    if (INVALID_SHEET_NAME_CHARS.test(name)) {
      throw new AuthoringError(
        `worksheet name "${name}" contains a character Excel forbids (* ? : \\ / [ ])`,
      );
    }
    if (name.startsWith("'") || name.endsWith("'")) {
      throw new AuthoringError(`worksheet name "${name}" cannot start or end with an apostrophe`);
    }
    if (this.getWorksheet(name) !== undefined) {
      throw new AuthoringError(
        `a worksheet named "${name}" already exists (names are case-insensitive)`,
      );
    }
  }

  /**
   * The codec's channel into this workbook — see `core/internal.ts` for why these are not public
   * methods. Declared last so every private field it closes over is already in scope.
   */
  readonly [INTERNAL]: WorkbookInternals = {
    addPreservedReference: (reference) => {
      this.#preservedReferences.push(reference);
    },
    addPreservedRootReference: (reference) => {
      this.#preservedRootReferences.push(reference);
    },
    restoreDifferentialStyles: (fragments) => {
      replaceContents(this.#dxfs, fragments);
    },
    restoreIndexedColors: (fragments) => {
      replaceContents(this.#indexedColors, fragments);
    },
    restoreMruColors: (fragments) => {
      replaceContents(this.#mruColors, fragments);
    },
    restoreTableStyles: (table) => {
      this.#tableStyles = table;
    },
    restoreThemePart: (theme) => {
      this.#theme = theme;
      this.#themeColors = undefined;
    },
    restoreNamedStyles: (styles) => {
      replaceContents(this.#namedStyles, styles);
    },
    restorePersons: (persons) => {
      this.#persons.clear();
      for (const person of persons) this.#persons.set(person.id, person);
    },
  };
}

/**
 * What a codec may do to a `Workbook` that an author may not: reinstate state read from a file, in
 * the finished form the file stated it. Reached as `workbook[INTERNAL]`; see `core/internal.ts`.
 *
 * Every operation replaces what it restores rather than merging, because a reader states a table
 * whole — a half-restored `<dxfs>` would leave existing `dxfId` references pointing into a mix of
 * two files.
 */
export interface WorkbookInternals {
  /** Record a workbook-level preserved reference (a pivot or slicer cache) read from a file. */
  addPreservedReference(reference: PreservedWorkbookReference): void;

  /** Record a package-root preserved reference (a customUI ribbon part, custom props) read from a file. */
  addPreservedRootReference(reference: PreservedRootReference): void;

  /**
   * Reinstate the differential-style (`<dxfs>`) table read from a file — the deserialization
   * counterpart the writer re-emits verbatim. Each entry is one `<dxf>…</dxf>` fragment, preserved as
   * opaque XML so a conditional-formatting rule's `dxfId` (an index into this table) stays valid on
   * re-write.
   */
  restoreDifferentialStyles(fragments: readonly string[]): void;

  /**
   * Reinstate the custom indexed-color palette (`<colors><indexedColors>`) read from a file — each
   * entry a verbatim `<rgbColor rgb="…"/>` fragment — so a colour referenced by `indexed="…"` keeps
   * its intended RGB on re-write instead of the palette being dropped and the colour shifting to a
   * default-palette entry.
   */
  restoreIndexedColors(fragments: readonly string[]): void;

  /**
   * Reinstate the most-recently-used colour swatches (`<colors><mruColors>`) read from a file, each
   * entry a verbatim `<color rgb="…"/>` fragment — the "Recent Colors" row a spreadsheet application
   * offers, which is the author's own working set rather than anything the model interprets.
   */
  restoreMruColors(fragments: readonly string[]): void;

  /**
   * Reinstate the custom table-style definitions (`<tableStyles>`) read from a file — see
   * {@link TableStyleTable} — so a table whose `styleName` names a custom style still resolves to a
   * real definition on re-write instead of dangling, and the file's nominated default table/pivot
   * styles survive.
   */
  restoreTableStyles(table: TableStyleTable): void;

  /**
   * Reinstate the theme part read from a file — opaque preserved XML plus the closure of parts it
   * reaches (see {@link PreservedTheme}) — so a workbook's colour and font schemes survive a re-write
   * instead of being replaced by the library's default Office theme. Passing `undefined` drops back
   * to that default.
   */
  restoreThemePart(theme: PreservedTheme | undefined): void;

  /**
   * Reinstate the named cell styles (`cellStyleXfs`/`cellStyles`) read from a file, index for index,
   * so a cell's link to a named style (its `xfId`) stays valid on re-write. Index 0 is the Normal
   * default.
   */
  restoreNamedStyles(styles: readonly NamedCellStyle[]): void;

  /**
   * Reinstate the threaded-comment identity registry (`xl/persons/person.xml`) read from a file — the
   * authors and mentioned people a comment thread's messages point at.
   *
   * Entries are keyed by {@link Person.id} and by nothing else. A single human legitimately owns
   * several entries: Excel interns a *mentioned* identity as its own `providerId="PeoplePicker"` entry
   * beside that person's `providerId="AD"` authoring entry — same `displayName`, same `userId`, a
   * different id — and points the mention at the new one. Collapsing entries by name or `userId` would
   * merge those two and silently re-point every mention at the wrong identity.
   *
   * {@link Workbook.addPerson} is the authoring verb.
   */
  restorePersons(persons: readonly Person[]): void;
}
