// Anchored images on the wire: the `xl/drawings/drawing{n}.xml` part (a DrawingML two-cell anchor
// per image), the drawing's own relationships to the `xl/media/` bytes, and the reader that turns a
// drawing back into anchors. The image bytes themselves are opaque here — the writer copies them
// verbatim into a media part and the reader hands them back untouched.

import {
  type AnchorPoint,
  type Extent,
  type ImageAnchor,
  type ImageEditAs,
  isOneCellAnchor,
} from '../../core/image.ts';
import {localName, parseXml} from './xml-read.ts';
import {XML_DECLARATION} from './xml.ts';

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const IMAGE_REL_TYPE = `${R_NS}/image`;

// The content type Excel expects for each image kind, keyed by lower-case extension. An unlisted
// extension falls back to `image/<ext>`, which is what a well-behaved consumer infers anyway.
const IMAGE_CONTENT_TYPES: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  svg: 'image/svg+xml',
};

/** The content type for a media part's `<Default Extension>` entry in `[Content_Types].xml`. */
export function imageContentType(extension: string): string {
  const ext = extension.toLowerCase();
  return IMAGE_CONTENT_TYPES[ext] ?? `image/${ext}`;
}

/** One image placed in a drawing: where it sits and the drawing-local relationship id that ties it
 * to its media bytes. */
export interface DrawingImage {
  readonly anchor: ImageAnchor;
  /** The `r:embed` id referencing this image's entry in the drawing's own `.rels`. */
  readonly embedId: string;
}

/** The `xl/drawings/drawing{n}.xml` part: one anchor per image, two-cell or one-cell by its shape. */
export function drawingXml(images: readonly DrawingImage[]): string {
  const anchors = images.map((image, i) => anchorXml(image, i + 1)).join('');
  return (
    XML_DECLARATION +
    `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
    anchors +
    '</xdr:wsDr>'
  );
}

function anchorXml(image: DrawingImage, id: number): string {
  const {anchor} = image;
  return isOneCellAnchor(anchor)
    ? oneCellAnchorXml(anchor.from, anchor.ext, anchor.rotation, image.embedId, id)
    : twoCellAnchorXml(anchor.from, anchor.to, anchor.editAs ?? 'oneCell', anchor.rotation, image.embedId, id);
}

// A picture anchored between two grid points. The geometry lives entirely in <xdr:from>/<xdr:to>, so
// the picture carries no absolute <a:xfrm> — a zeroed one would override the anchor and collapse the
// image to nothing in strict viewers (LibreOffice), while a non-zero one would fight the anchor. A
// rotation is the one transform kept: it can't be derived from the anchor, so it rides a rot-only xfrm.
function twoCellAnchorXml(
  from: AnchorPoint,
  to: AnchorPoint,
  editAs: ImageEditAs,
  rotation: number | undefined,
  embedId: string,
  id: number
): string {
  return (
    `<xdr:twoCellAnchor editAs="${editAs}">` +
    `<xdr:from>${anchorPointXml(from)}</xdr:from>` +
    `<xdr:to>${anchorPointXml(to)}</xdr:to>` +
    picXml(embedId, id, rotation) +
    '<xdr:clientData/>' +
    '</xdr:twoCellAnchor>'
  );
}

// A picture pinned at one grid point with a fixed EMU extent. editAs is a two-cell-only attribute and
// the schema forbids it here, so a one-cell anchor never carries one.
function oneCellAnchorXml(
  from: AnchorPoint,
  ext: Extent,
  rotation: number | undefined,
  embedId: string,
  id: number
): string {
  return (
    '<xdr:oneCellAnchor>' +
    `<xdr:from>${anchorPointXml(from)}</xdr:from>` +
    `<xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>` +
    picXml(embedId, id, rotation) +
    '<xdr:clientData/>' +
    '</xdr:oneCellAnchor>'
  );
}

function picXml(embedId: string, id: number, rotation: number | undefined): string {
  const xfrm = rotation !== undefined ? `<a:xfrm rot="${rotation}"/>` : '';
  return (
    '<xdr:pic>' +
    `<xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/>` +
    '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
    `<xdr:blipFill><a:blip r:embed="${embedId}"/>` +
    '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    `<xdr:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>` +
    '</xdr:pic>'
  );
}

function anchorPointXml(point: AnchorPoint): string {
  return (
    `<xdr:col>${point.col}</xdr:col><xdr:colOff>${point.colOff ?? 0}</xdr:colOff>` +
    `<xdr:row>${point.row}</xdr:row><xdr:rowOff>${point.rowOff ?? 0}</xdr:rowOff>`
  );
}

/** The drawing's `_rels/drawing{n}.xml.rels`: one image relationship per anchor, in `embedId` order
 * (`rId1`, `rId2`, …), each pointing at the media part the anchor shows. */
export function drawingRelsXml(mediaTargets: readonly string[]): string {
  const rels = mediaTargets
    .map((target, i) => `<Relationship Id="rId${i + 1}" Type="${IMAGE_REL_TYPE}" Target="${target}"/>`)
    .join('');
  return XML_DECLARATION + `<Relationships xmlns="${PKG_RELS_NS}">${rels}</Relationships>`;
}

/** An image anchor parsed from a drawing part, with the `r:embed` id that names its media. A two-cell
 * anchor carries `to` (and may carry `editAs`); a one-cell anchor carries `ext` instead. */
export interface ParsedImageAnchor {
  readonly from: AnchorPoint;
  readonly to?: AnchorPoint;
  readonly ext?: Extent;
  readonly editAs?: ImageEditAs;
  readonly rotation?: number;
  readonly embed: string;
}

type PointDraft = {col: number; row: number; colOff: number; rowOff: number};

function blankPoint(): PointDraft {
  return {col: 0, row: 0, colOff: 0, rowOff: 0};
}

const EDIT_AS = new Set<string>(['oneCell', 'twoCell', 'absolute']);

/** Parse a drawing part into its image anchors (both `<xdr:twoCellAnchor>` and `<xdr:oneCellAnchor>`).
 * Anchors that are not pictures (a chart, a shape) carry no `<a:blip r:embed>` and are skipped, so a
 * mixed drawing yields only its images. */
export function parseDrawing(xml: string): ParsedImageAnchor[] {
  const anchors: ParsedImageAnchor[] = [];
  let from: PointDraft | null = null;
  let to: PointDraft | null = null;
  let ext: Extent | undefined;
  let editAs: ImageEditAs | undefined;
  let rotation: number | undefined;
  let embed: string | undefined;
  // The point (<xdr:from> or <xdr:to>) whose coordinate children are currently streaming in.
  let target: PointDraft | null = null;
  // Depth inside <xdr:pic>, so the anchor-level <xdr:ext> is not confused with the <a:ext> nested in
  // a picture's spPr transform (both have local name "ext").
  let picDepth = 0;
  // Which coordinate child is open, so its text lands on the right field; '' between children.
  let coord = '';
  let text = '';

  parseXml(xml, {
    onOpen(name, attrs) {
      const local = localName(name);
      if (local === 'twoCellAnchor' || local === 'oneCellAnchor') {
        from = blankPoint();
        to = local === 'twoCellAnchor' ? blankPoint() : null;
        ext = undefined;
        rotation = undefined;
        embed = undefined;
        const mode = attrs.editAs;
        editAs = mode !== undefined && EDIT_AS.has(mode) ? (mode as ImageEditAs) : undefined;
      } else if (local === 'pic') {
        picDepth++;
      } else if (local === 'xfrm' && picDepth > 0) {
        // The picture's own rotation — the one spPr transform that can't be derived from the anchor.
        const rot = Number(attrs.rot);
        if (Number.isFinite(rot) && rot !== 0) rotation = rot;
      } else if (local === 'from') {
        target = from;
      } else if (local === 'to') {
        target = to;
      } else if (local === 'ext' && picDepth === 0) {
        const cx = Number(attrs.cx);
        const cy = Number(attrs.cy);
        if (Number.isFinite(cx) && Number.isFinite(cy)) ext = {cx, cy};
      } else if (local === 'blip') {
        const value = attrs['r:embed'] ?? attrs.embed;
        if (value !== undefined) embed = value;
      } else if (target !== null && COORDINATES.has(local)) {
        coord = local;
        text = '';
      }
    },
    onText(chunk) {
      if (coord !== '') text += chunk;
    },
    onClose(name) {
      const local = localName(name);
      if (target !== null && coord === local && COORDINATES.has(local)) {
        const value = Number(text);
        if (Number.isFinite(value)) setCoordinate(target, local, value);
        coord = '';
      } else if (local === 'from' || local === 'to') {
        target = null;
      } else if (local === 'pic') {
        picDepth--;
      } else if (local === 'twoCellAnchor' || local === 'oneCellAnchor') {
        if (from !== null && embed !== undefined) {
          const rot = rotation !== undefined ? {rotation} : {};
          if (to !== null) {
            const mode = editAs !== undefined ? {editAs} : {};
            anchors.push({from: {...from}, to: {...to}, ...mode, ...rot, embed});
          } else if (ext !== undefined) {
            anchors.push({from: {...from}, ext, ...rot, embed});
          }
        }
        from = null;
        to = null;
        ext = undefined;
      }
    },
  });
  return anchors;
}

// Anchor content a drawing can hold that the image model does not interpret: a chart
// (`<xdr:graphicFrame>`), a shape or text box (`<xdr:sp>`), a connector (`<xdr:cxnSp>`), or a group
// (`<xdr:grpSp>`). A drawing carrying any of these is preserved whole rather than modeled, so it is
// not re-serialised from its pictures alone (which would silently drop the chart/shape).
const UNMODELED_DRAWING_CONTENT = new Set<string>(['graphicFrame', 'sp', 'cxnSp', 'grpSp']);

/** Whether a drawing part holds anchor content beyond plain pictures — a chart, shape, connector, or
 * group. Excel packs every one of a sheet's anchors into a single drawing part, so a sheet with both a
 * picture and a chart yields a mixed drawing; modeling only its pictures and re-serialising from them
 * would drop the chart. The reader uses this to fall back to whole-drawing byte-preservation instead. */
export function drawingHasUnmodeledContent(xml: string): boolean {
  let found = false;
  parseXml(xml, {
    onOpen(name) {
      if (!found && UNMODELED_DRAWING_CONTENT.has(localName(name))) found = true;
    },
    onText() {},
    onClose() {},
  });
  return found;
}

const COORDINATES = new Set<string>(['col', 'colOff', 'row', 'rowOff']);

function setCoordinate(point: PointDraft, coord: string, value: number): void {
  if (coord === 'col') point.col = value;
  else if (coord === 'colOff') point.colOff = value;
  else if (coord === 'row') point.row = value;
  else point.rowOff = value;
}
