// Anchored images: the model for a picture pinned to a worksheet's grid.
//
// The image *bytes* live once on the workbook (a small media registry, addressed by a numeric id);
// a worksheet then anchors that image to a rectangle of cells. Storing the bytes centrally means the
// same picture used on two sheets — a logo in a header band, say — is one media part, not two.

/** A point in the drawing grid: a 0-based column and row, plus an EMU offset into that cell.
 * The offsets default to zero, pinning the point to the cell's top-left corner. */
export interface AnchorPoint {
  /** 0-based column index (column A is 0). */
  readonly col: number;
  /** 0-based row index (row 1 is 0). */
  readonly row: number;
  /** Horizontal offset into the cell, in EMUs (914400 per inch). Defaults to 0. */
  readonly colOff?: number;
  /** Vertical offset into the cell, in EMUs. Defaults to 0. */
  readonly rowOff?: number;
}

/** EMUs per pixel at Excel's notional 96 DPI (914400 EMU/inch ÷ 96 px/inch). The conversion is
 * DPI-independent by construction: a pixel extent is a fixed physical size regardless of screen. */
export const PX_TO_EMU = 9525;

/** How a two-cell-anchored image tracks edits to the cells it spans. `twoCell` moves and resizes with
 * them; `oneCell` moves but keeps its size; `absolute` is pinned to the page and does neither. Excel
 * defaults to `oneCell` when the attribute is omitted. */
export type ImageEditAs = 'oneCell' | 'twoCell' | 'absolute';

/** A fixed image size in EMUs — the extent of a one-cell anchor, which pixel dimensions convert into
 * via {@link PX_TO_EMU}. */
export interface Extent {
  readonly cx: number;
  readonly cy: number;
}

/** A two-cell anchor: the image's top-left (`from`) and bottom-right (`to`) grid points. The image
 * fills the rectangle between them and reflows as the intervening rows/columns resize; `editAs`
 * selects how strictly it follows. */
export interface TwoCellAnchor {
  readonly from: AnchorPoint;
  readonly to: AnchorPoint;
  readonly editAs?: ImageEditAs;
  /** Clockwise rotation in 1/60000 of a degree (`2700000` = 45°), preserved from a loaded file. */
  readonly rotation?: number;
}

/** A one-cell anchor: a single top-left grid point plus a fixed extent. The image keeps its size as
 * the grid resizes, moving only with its anchor cell. `editAs` is a two-cell-only attribute and has
 * no place here. */
export interface OneCellAnchor {
  readonly from: AnchorPoint;
  readonly ext: Extent;
  /** Clockwise rotation in 1/60000 of a degree (`2700000` = 45°), preserved from a loaded file. */
  readonly rotation?: number;
}

/** Where an image sits on the grid: a rectangle between two cells, or a point plus a fixed extent. */
export type ImageAnchor = TwoCellAnchor | OneCellAnchor;

/** Narrow an anchor to its one-cell (fixed-extent) form; the complement is {@link TwoCellAnchor}. */
export function isOneCellAnchor(anchor: ImageAnchor): anchor is OneCellAnchor {
  return 'ext' in anchor;
}

/** An image pinned to a worksheet: which workbook media it shows (`imageId`) and where. */
export interface AnchoredImage {
  /** Index into the workbook's media registry (the id {@link Workbook.addImage} returned). */
  readonly imageId: number;
  readonly anchor: ImageAnchor;
}

/** A picture's bytes and its file kind, as held in the workbook's media registry. */
export interface WorkbookImage {
  /** Lower-case file extension without a dot — `"png"`, `"jpeg"`, `"gif"`. Drives the media part's
   * name and content type. */
  readonly extension: string;
  readonly data: Uint8Array;
}

// Leading magic bytes for the raster formats a spreadsheet embeds, most-specific first. Used to infer
// an extension when the caller supplies none, so a package never declares an `image/undefined` type.
const IMAGE_MAGIC: ReadonlyArray<{readonly ext: string; readonly sig: readonly number[]}> = [
  {ext: 'png', sig: [0x89, 0x50, 0x4e, 0x47]},
  {ext: 'jpeg', sig: [0xff, 0xd8, 0xff]},
  {ext: 'gif', sig: [0x47, 0x49, 0x46]},
  {ext: 'bmp', sig: [0x42, 0x4d]},
  {ext: 'tiff', sig: [0x49, 0x49, 0x2a, 0x00]},
  {ext: 'tiff', sig: [0x4d, 0x4d, 0x00, 0x2a]},
];

function sniffImageExtension(data: Uint8Array): string {
  for (const {ext, sig} of IMAGE_MAGIC) {
    if (sig.every((b, i) => data[i] === b)) return ext;
  }
  // An unrecognised blob still needs a valid media name and content type; png keeps the package
  // well-formed rather than emitting a `<Default>` with no or a bogus extension.
  return 'png';
}

/** Reduce a caller-supplied extension to the bare, lower-case alphanumeric token OOXML expects for a
 * media part's name and `<Default Extension>`. A leading dot (`".png"`), a URL query string
 * (`"png?alt=media"`), or any other separator a real-world filename/URL drags in is stripped to the
 * leading run of alphanumerics; a missing or all-punctuation hint falls back to sniffing the bytes'
 * magic number, so the package is always well-formed. */
export function normalizeImageExtension(extension: string | undefined, data: Uint8Array): string {
  if (typeof extension === 'string') {
    const token = extension.toLowerCase().match(/[a-z0-9]+/)?.[0];
    if (token !== undefined) return token;
  }
  return sniffImageExtension(data);
}
