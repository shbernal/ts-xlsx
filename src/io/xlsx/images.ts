// Anchored images on the wire: the `xl/drawings/drawing{n}.xml` part (a DrawingML two-cell anchor
// per image), the drawing's own relationships to the `xl/media/` bytes, and the reader that turns a
// drawing back into anchors. The image bytes themselves are opaque here: the writer copies them
// verbatim into a media part and the reader hands them back untouched.

import {
  type AnchorPoint,
  type Extent,
  type ImageAnchor,
  type ImageEditAs,
  isImageEditAs,
  isOneCellAnchor,
} from '../../core/image.ts';
import {enumToken, localName, numFinite, parseXml, TextCapture} from '../../xml/xml-read.ts';
import {checkedToken, numAttr, numberText, XML_DECLARATION} from '../../xml/xml.ts';
import {RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {relationship, relationshipsPart} from '../opc/rels.ts';
import {DRAWINGML_NS, XDR_NS} from './namespaces.ts';

const IMAGE_REL_TYPE = `${RELATIONSHIPS_NS}/image`;

// The content type Excel expects for each image kind, keyed by lower-case extension. An unlisted
// extension falls back to `image/<ext>`, which is what a well-behaved consumer infers anyway.
// A Map, not an object literal: the extension comes off a media part's name in the package, and an
// object would answer `constructor` with a function that then stringifies into a content-type
// attribute.
const IMAGE_CONTENT_TYPES: ReadonlyMap<string, string> = new Map([
  ['png', 'image/png'],
  ['jpg', 'image/jpeg'],
  ['jpeg', 'image/jpeg'],
  ['gif', 'image/gif'],
  ['bmp', 'image/bmp'],
  ['tif', 'image/tiff'],
  ['tiff', 'image/tiff'],
  ['emf', 'image/x-emf'],
  ['wmf', 'image/x-wmf'],
  ['svg', 'image/svg+xml'],
]);

/** The content type for a media part's `<Default Extension>` entry in `[Content_Types].xml`. */
export function imageContentType(extension: string): string {
  const ext = extension.toLowerCase();
  return IMAGE_CONTENT_TYPES.get(ext) ?? `image/${ext}`;
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
    `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="${RELATIONSHIPS_NS}">` +
    anchors +
    '</xdr:wsDr>'
  );
}

function anchorXml(image: DrawingImage, id: number): string {
  const {anchor} = image;
  return isOneCellAnchor(anchor)
    ? oneCellAnchorXml(anchor.from, anchor.ext, anchor.rotation, image.embedId, id)
    : twoCellAnchorXml(
        anchor.from,
        anchor.to,
        anchor.editAs ?? 'oneCell',
        anchor.rotation,
        image.embedId,
        id,
      );
}

// A picture anchored between two grid points. The geometry lives entirely in <xdr:from>/<xdr:to>, so
// the picture carries no absolute <a:xfrm>: a zeroed one would override the anchor and collapse the
// image to nothing in strict viewers (LibreOffice), while a non-zero one would fight the anchor. A
// rotation is the one transform kept: it can't be derived from the anchor, so it rides a rot-only xfrm.
function twoCellAnchorXml(
  from: AnchorPoint,
  to: AnchorPoint,
  editAs: ImageEditAs,
  rotation: number | undefined,
  embedId: string,
  id: number,
): string {
  return (
    `<xdr:twoCellAnchor editAs="${checkedToken(editAs, isImageEditAs, 'image anchor edit mode')}">` +
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
  id: number,
): string {
  return (
    '<xdr:oneCellAnchor>' +
    `<xdr:from>${anchorPointXml(from)}</xdr:from>` +
    `<xdr:ext${numAttr('cx', ext.cx)}${numAttr('cy', ext.cy)}/>` +
    picXml(embedId, id, rotation) +
    '<xdr:clientData/>' +
    '</xdr:oneCellAnchor>'
  );
}

function picXml(embedId: string, id: number, rotation: number | undefined): string {
  const xfrm = rotation !== undefined ? `<a:xfrm${numAttr('rot', rotation)}/>` : '';
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

// The grid point's four numbers are all author-reachable through `addImageAnchor`, and each is an
// `xsd:int` or an EMU offset with no spelling for a non-finite value, so they are refused on the same
// terms as the extent above. That they are elements rather than attributes changes nothing.
function anchorPointXml(point: AnchorPoint): string {
  return (
    `<xdr:col>${numberText(point.col)}</xdr:col>` +
    `<xdr:colOff>${numberText(point.colOff ?? 0)}</xdr:colOff>` +
    `<xdr:row>${numberText(point.row)}</xdr:row>` +
    `<xdr:rowOff>${numberText(point.rowOff ?? 0)}</xdr:rowOff>`
  );
}

/** The drawing's `_rels/drawing{n}.xml.rels`: one image relationship per anchor, in `embedId` order
 * (`rId1`, `rId2`, …), each pointing at the media part the anchor shows. */
export function drawingRelsXml(mediaTargets: readonly string[]): string {
  return relationshipsPart(
    mediaTargets.map((target, i) => relationship(`rId${i + 1}`, IMAGE_REL_TYPE, target)),
  );
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
  // The coordinate child currently streaming in, so its text lands on the right field.
  const coord = new TextCapture(COORDINATES);

  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      const local = localName(name);
      if (local === 'twoCellAnchor' || local === 'oneCellAnchor') {
        from = blankPoint();
        to = local === 'twoCellAnchor' ? blankPoint() : null;
        ext = undefined;
        rotation = undefined;
        embed = undefined;
        editAs = enumToken(attrs.editAs, isImageEditAs);
      } else if (local === 'pic') {
        picDepth++;
      } else if (local === 'xfrm' && picDepth > 0) {
        // The picture's own rotation: the one spPr transform that can't be derived from the anchor.
        const rot = numFinite(attrs.rot);
        if (rot !== undefined && rot !== 0) rotation = rot;
      } else if (local === 'from') {
        target = from;
      } else if (local === 'to') {
        target = to;
      } else if (local === 'ext' && picDepth === 0) {
        const cx = numFinite(attrs.cx, 0);
        const cy = numFinite(attrs.cy, 0);
        if (cx !== undefined && cy !== undefined) ext = {cx, cy};
      } else if (local === 'blip') {
        const value = attrs['r:embed'] ?? attrs.embed;
        if (value !== undefined) embed = value;
      } else if (target !== null) {
        coord.open(local, selfClosing);
      }
    },
    onText(chunk) {
      coord.text(chunk);
    },
    onClose(name) {
      const local = localName(name);
      const text = coord.close(local);
      if (text !== undefined) {
        const value = numFinite(text);
        if (target !== null && value !== undefined) setCoordinate(target, local, value);
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

/** Whether a drawing part holds anchor content beyond plain pictures: a chart, shape, connector, or
 * group. Excel packs every one of a sheet's anchors into a single drawing part, so a sheet with both a
 * picture and a chart yields a mixed drawing; modeling only its pictures and re-serialising from them
 * would drop the chart. The reader uses this to fall back to whole-drawing byte-preservation instead. */
export function drawingHasUnmodeledContent(xml: string): boolean {
  let found = false;
  parseXml(xml, {
    onOpen(name) {
      if (!found && UNMODELED_DRAWING_CONTENT.has(localName(name))) found = true;
    },
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
