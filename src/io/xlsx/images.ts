// Anchored images on the wire: the `xl/drawings/drawing{n}.xml` part (a DrawingML two-cell anchor
// per image), the drawing's own relationships to the `xl/media/` bytes, and the reader that turns a
// drawing back into anchors. The image bytes themselves are opaque here — the writer copies them
// verbatim into a media part and the reader hands them back untouched.

import type {AnchorPoint, ImageAnchor} from '../../core/image.ts';
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

/** The `xl/drawings/drawing{n}.xml` part: one two-cell anchor per image. */
export function drawingXml(images: readonly DrawingImage[]): string {
  const anchors = images.map((image, i) => twoCellAnchorXml(image, i + 1)).join('');
  return (
    XML_DECLARATION +
    `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
    anchors +
    '</xdr:wsDr>'
  );
}

// A picture anchored between two grid points. The geometry lives entirely in <xdr:from>/<xdr:to>;
// the <a:xfrm> offset/extent are placeholders Excel recomputes from the anchor, so they stay zero.
function twoCellAnchorXml(image: DrawingImage, id: number): string {
  const {from, to} = image.anchor;
  return (
    '<xdr:twoCellAnchor editAs="oneCell">' +
    `<xdr:from>${anchorPointXml(from)}</xdr:from>` +
    `<xdr:to>${anchorPointXml(to)}</xdr:to>` +
    '<xdr:pic>' +
    `<xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/>` +
    '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
    `<xdr:blipFill><a:blip r:embed="${image.embedId}"/>` +
    '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
    '</xdr:pic>' +
    '<xdr:clientData/>' +
    '</xdr:twoCellAnchor>'
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

/** A two-cell anchor parsed from a drawing part, with the `r:embed` id that names its media. */
export interface ParsedImageAnchor {
  readonly from: AnchorPoint;
  readonly to: AnchorPoint;
  readonly embed: string;
}

type PointDraft = {col: number; row: number; colOff: number; rowOff: number};

function blankPoint(): PointDraft {
  return {col: 0, row: 0, colOff: 0, rowOff: 0};
}

/** Parse a drawing part into its two-cell image anchors. Anchors that are not pictures (a chart, a
 * shape) carry no `<a:blip r:embed>` and are skipped, so a mixed drawing yields only its images. */
export function parseDrawing(xml: string): ParsedImageAnchor[] {
  const anchors: ParsedImageAnchor[] = [];
  let from: PointDraft | null = null;
  let to: PointDraft | null = null;
  let embed: string | undefined;
  // The point (<xdr:from> or <xdr:to>) whose coordinate children are currently streaming in.
  let target: PointDraft | null = null;
  // Which coordinate child is open, so its text lands on the right field; '' between children.
  let coord = '';
  let text = '';

  parseXml(xml, {
    onOpen(name, attrs) {
      const local = localName(name);
      if (local === 'twoCellAnchor') {
        from = blankPoint();
        to = blankPoint();
        embed = undefined;
      } else if (local === 'from') {
        target = from;
      } else if (local === 'to') {
        target = to;
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
      } else if (local === 'twoCellAnchor') {
        if (from !== null && to !== null && embed !== undefined) {
          anchors.push({from: {...from}, to: {...to}, embed});
        }
        from = null;
        to = null;
      }
    },
  });
  return anchors;
}

const COORDINATES = new Set<string>(['col', 'colOff', 'row', 'rowOff']);

function setCoordinate(point: PointDraft, coord: string, value: number): void {
  if (coord === 'col') point.col = value;
  else if (coord === 'colOff') point.colOff = value;
  else if (coord === 'row') point.row = value;
  else point.rowOff = value;
}
