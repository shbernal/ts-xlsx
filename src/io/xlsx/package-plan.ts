// The plan layer of the writer: pure graph resolution that turns a Workbook model into the numbered,
// cross-referenced set of parts an `.xlsx` package needs — media, preserved (verbatim-carried) parts,
// and the sheet-/workbook-local relationship ids that wire them — before any XML is serialised.

import type {PivotTable} from '../../core/pivot-table.ts';
import type {Table} from '../../core/table.ts';
import type {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import type {NoteCell} from './comments.ts';
import type {DrawingImage} from './images.ts';
import {extensionOf, relativePartPath, relsPathForPart} from './part-paths.ts';
import {preservedRelsXml} from './relationships.ts';

// A sheet's relationship-id allocator: hands out `rId1`, `rId2`, … in the one canonical order the
// package wires a sheet's parts (tables, drawing, comments, printer settings, external hyperlinks,
// background, preserved references, pivot tables). Every sheet-local id is drawn from here in
// sequence, so no plan step re-derives its starting offset by summing the counts of the steps before
// it — the arithmetic that, open-coded once per step with subtly different prefixes, could silently
// hand two parts the same id and corrupt the package. Monotonic by construction, so collisions cannot
// arise however the steps grow. One fresh allocator per sheet; the ids it yields are sheet-local.
export class SheetRelIds {
  #next = 1;
  /** The next relationship id (`rId1`, `rId2`, …), advancing the counter. */
  next(): string {
    return `rId${this.#next++}`;
  }
}

// A pivot table planned for emission: its global part number, the workbook-unique `cacheId` its
// `<pivotCaches>` registration and `pivotTableDefinition` agree on, the sheet-local relationship
// linking its host sheet to the pivot-table part, and the workbook relationship reaching its cache
// definition (assigned once the modeled workbook rels are counted).
export interface PivotPlan {
  readonly number: number;
  readonly cacheId: string;
  readonly table: PivotTable;
  readonly sheetRelId: string;
  workbookRelId: string;
}

// A sheet's notes paired with the part number and sheet-local relationship ids that link the sheet
// to its comments part (by type) and its VML drawing (by the `<legacyDrawing>` element).
export interface CommentPlan {
  readonly number: number;
  readonly notes: readonly NoteCell[];
  readonly vmlRelId: string;
  readonly commentsRelId: string;
}

// A sheet's opaque printer-settings blob paired with the part number naming its `.bin` part and the
// sheet-local relationship id that links the sheet's `<pageSetup r:id>` to it.
export interface PrinterSettingsPlan {
  readonly number: number;
  readonly data: Uint8Array;
  readonly relId: string;
}

// A table paired with the identifiers the package needs: a workbook-global part number
// and the sheet-local relationship id that links its worksheet to the table part.
export interface PlannedTable {
  readonly table: Table;
  readonly number: number;
  readonly relId: string;
}

// A sheet background image resolved for serialisation: the sheet-local relationship id its
// `<picture>` element references, and the media part (global number + extension) that holds the bytes.
export interface BackgroundPlan {
  readonly relId: string;
  readonly mediaNumber: number;
  readonly extension: string;
}

// A verbatim-preserved package part resolved for serialisation: the collision-proof path it is
// emitted at, its bytes and content type, and — when it references other parts — the rels part
// linking it to their new paths.
export interface PreservedPartPlan {
  readonly path: string;
  readonly bytes: Uint8Array;
  readonly contentType: string;
  readonly relsPath: string | null;
  readonly relsXml: string | null;
}

// A preserved worksheet reference resolved for serialisation, short of its sheet-local relationship
// id: the worksheet element that wires it (`null` for a pivot-table/slicer reference the sheet carries
// by relationship alone), the relationship Type, and the new path of the entry part it targets. The id
// is assigned by the caller from the sheet's {@link SheetRelIds} allocator, at the reference's
// canonical position in the sheet-local id sequence (after tables/drawing/comments/printer-settings/
// external-hyperlinks/background) — so a preserved reference never renumbers an id already threaded
// into the sheet XML.
export interface ResolvedPreservedReference {
  readonly element: 'drawing' | 'legacyDrawingHF' | null;
  readonly relType: string;
  readonly entryPath: string;
}

// A resolved preserved reference with its sheet-local relationship id assigned.
export interface PreservedReferencePlan extends ResolvedPreservedReference {
  readonly relId: string;
}

// A preserved workbook reference resolved for serialisation: its relationship Type, the new path of
// the entry part, and — for a pivot cache — the `cacheId` its `<pivotCaches>` registration carries.
// The workbook relationship id is assigned at emit time (it follows the modeled workbook rels, whose
// count depends on whether a shared-strings part is emitted), so it is not fixed here.
export interface PreservedWorkbookReferencePlan {
  readonly relType: string;
  readonly entryPath: string;
  readonly pivotCacheId: string | undefined;
}

// The whole workbook's preserved content resolved for serialisation: per-sheet references (parallel to
// the sheets) still awaiting their sheet-local relationship ids, the workbook-level reference plans,
// and the flat, de-duplicated list of parts to emit. Kept together because the parts are numbered
// globally while the references are sheet- or workbook-local.
export interface PreservedPlan {
  readonly perSheet: readonly (readonly ResolvedPreservedReference[])[];
  readonly workbook: readonly PreservedWorkbookReferencePlan[];
  readonly parts: readonly PreservedPartPlan[];
}

// A sheet's drawing part: its workbook-global number, the sheet-local relationship id linking the
// sheet's `<drawing>` element to it, and the images it lays out.
export interface DrawingPlan {
  readonly number: number;
  readonly relId: string;
  readonly images: readonly PlannedImage[];
}

// An anchored image resolved for serialisation: its anchor and drawing-local embed id (via
// DrawingImage) plus the media part number and extension its embed relationship targets.
export interface PlannedImage extends DrawingImage {
  readonly mediaNumber: number;
  readonly extension: string;
}

// One picture written to `xl/media/`: its global part number, extension, and bytes.
export interface MediaPart {
  readonly number: number;
  readonly extension: string;
  readonly data: Uint8Array;
}

// The workbook's media, resolved for writing: the parts to emit, a map from a workbook image id to
// its media part number (so a drawing embed can target it), and the distinct extensions in use (so
// content types can declare an image `<Default>` per extension).
export interface MediaPlan {
  readonly parts: readonly MediaPart[];
  readonly numberById: ReadonlyMap<number, number>;
  readonly extensions: readonly string[];
}

// Per-kind counters for numbering preserved parts, each seeded past the generated parts of its kind.
interface PreservedNumbering {
  drawing: number;
  vml: number;
  media: number;
}

// Gather the workbook images actually referenced by some sheet — either anchored in a drawing or set
// as a sheet background (an unreferenced image is not written) — number them in first-use order, and
// record the extensions in play. A sheet referencing an id with no registered image is a programming
// error the writer surfaces rather than emitting a dangling relationship.
export function planMedia(workbook: Workbook, sheets: readonly Worksheet[]): MediaPlan {
  const usedIds: number[] = [];
  const seen = new Set<number>();
  const use = (id: number): void => {
    if (!seen.has(id)) {
      seen.add(id);
      usedIds.push(id);
    }
  };
  for (const sheet of sheets) {
    for (const image of sheet.images) use(image.imageId);
    if (sheet.backgroundImageId !== undefined) use(sheet.backgroundImageId);
  }
  const parts: MediaPart[] = [];
  const numberById = new Map<number, number>();
  const extensions = new Set<string>();
  usedIds.forEach((id, i) => {
    const image = workbook.getImage(id);
    if (image === undefined) {
      throw new Error(
        `a worksheet anchors image id ${id}, which is not registered on the workbook`,
      );
    }
    const number = i + 1;
    parts.push({number, extension: image.extension, data: image.data});
    numberById.set(id, number);
    extensions.add(image.extension);
  });
  return {parts, numberById, extensions: [...extensions]};
}

// Resolve every sheet's verbatim-preserved worksheet references (a vector-shape drawing, a
// header/footer image) into the parts to emit and the per-sheet reference data that wires them. Each
// reference's captured part closure is re-numbered onto collision-proof `preservedP{n}` paths — so
// preserved content never clobbers a generated drawing/VML/media part — with the closure's internal
// relationships rewritten to the new sibling paths. Part numbering is the only cross-sheet concern
// here; each reference's sheet-local relationship id is assigned by the caller from the sheet's
// {@link SheetRelIds} allocator, so this function stays free of the sheet-local id arithmetic.
export function planPreservedParts(
  workbook: Workbook,
  generatedDrawingCount: number,
  generatedMediaCount: number,
): PreservedPlan {
  const sheets = workbook.worksheets;
  // The writer generates drawings, VML, and media of its own, so a preserved part of one of those
  // kinds is re-numbered past the generated ones (a preserved drawing never clobbers an anchored
  // drawing, a preserved VML never clobbers a comment's VML). Comment VML is numbered by sheet index,
  // so `sheets.length` bounds it. Every other kind (pivot tables, caches, slicers, charts) the writer
  // never generates, so those keep their original path — see {@link preservedPartPath}.
  const numbering: PreservedNumbering = {
    drawing: generatedDrawingCount,
    vml: sheets.length,
    media: generatedMediaCount,
  };

  // One package-wide remap and one emitted-parts map: a part reached through more than one reference
  // (a pivot cache reached both from its pivot table and from the workbook) is numbered once and
  // emitted once, so overlapping closures collapse instead of duplicating parts.
  const remap = new Map<string, string>();
  const allReferences = [
    ...sheets.flatMap((sheet) => sheet.preservedReferences),
    ...workbook.preservedReferences,
  ];
  for (const reference of allReferences) {
    for (const part of reference.parts) {
      if (!remap.has(part.path)) remap.set(part.path, preservedPartPath(part.path, numbering));
    }
  }
  const emitted = new Map<string, PreservedPartPlan>();
  for (const reference of allReferences) {
    for (const part of reference.parts) {
      const newPath = remap.get(part.path) as string;
      if (emitted.has(newPath)) continue;
      const rels = part.rels.flatMap((rel) => {
        const target = remap.get(rel.targetPath);
        return target === undefined
          ? []
          : [{id: rel.id, type: rel.type, target: relativePartPath(newPath, target)}];
      });
      emitted.set(newPath, {
        path: newPath,
        bytes: part.bytes,
        contentType: part.contentType,
        relsPath: rels.length === 0 ? null : relsPathForPart(newPath),
        relsXml: rels.length === 0 ? null : preservedRelsXml(rels),
      });
    }
  }

  const perSheet = sheets.map((sheet): ResolvedPreservedReference[] =>
    sheet.preservedReferences.map((reference) => ({
      element: reference.element,
      relType: reference.relType,
      entryPath: remap.get(reference.entryPath) as string,
    })),
  );

  const workbookRefs = workbook.preservedReferences.map(
    (reference): PreservedWorkbookReferencePlan => ({
      relType: reference.relType,
      entryPath: remap.get(reference.entryPath) as string,
      pivotCacheId: reference.pivotCacheId,
    }),
  );

  return {perSheet, workbook: workbookRefs, parts: [...emitted.values()]};
}

// The path a preserved part is emitted at. A kind the writer generates of its own — a drawing, a VML,
// a media image — is re-numbered past the generated parts of that kind (see {@link planPreservedParts})
// so it never clobbers one. Every other kind (a pivot table, a pivot/slicer cache, a slicer, a chart)
// the writer never generates, so it keeps its original path — leaving the package's standard part
// names intact and letting overlapping closures agree on a single path for a shared part.
function preservedPartPath(originalPath: string, numbering: PreservedNumbering): string {
  const ext = extensionOf(originalPath);
  if (ext.toLowerCase() === 'vml') return `xl/drawings/vmlDrawing${++numbering.vml}.vml`;
  if (originalPath.startsWith('xl/media/')) return `xl/media/image${++numbering.media}.${ext}`;
  if (originalPath.startsWith('xl/drawings/') && ext.toLowerCase() === 'xml') {
    return `xl/drawings/drawing${++numbering.drawing}.xml`;
  }
  return originalPath;
}
