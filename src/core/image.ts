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
}

/** A one-cell anchor: a single top-left grid point plus a fixed extent. The image keeps its size as
 * the grid resizes, moving only with its anchor cell. `editAs` is a two-cell-only attribute and has
 * no place here. */
export interface OneCellAnchor {
  readonly from: AnchorPoint;
  readonly ext: Extent;
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
