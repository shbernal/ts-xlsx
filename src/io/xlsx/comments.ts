// Cell notes (legacy comments) — the `xl/comments{n}.xml` part, its `xl/drawings/vmlDrawing{n}.vml`
// companion, and the reader that maps a note back onto its cell.
//
// A note is anchored to a cell by A1 reference and rendered by Excel as a floating yellow box. The
// box's geometry lives in a legacy VML drawing (the pre-DrawingML shape format Excel still requires
// for notes); the text lives in the comments part. Both are emitted together — a comments part with
// no matching `<legacyDrawing>`/VML reads as text but renders nothing, so we never split them.

import type {Worksheet} from '../../core/worksheet.ts';
import {SPREADSHEETML_NS} from './namespaces.ts';
import {escapeText, needsSpacePreserve, XML_DECLARATION} from './xml.ts';
import {localName, parseXml} from './xml-read.ts';

/** A cell carrying a note, paired with the coordinates the VML anchor needs. */
export interface NoteCell {
  readonly ref: string;
  /** 1-based row of the noted cell. */
  readonly row: number;
  /** 1-based column of the noted cell. */
  readonly col: number;
  readonly text: string;
}

/** Gather every noted cell on a sheet. A note anchors to its cell regardless of the cell's value,
 * so a note on an otherwise-empty cell is collected too. */
export function collectNotes(sheet: Worksheet): NoteCell[] {
  const notes: NoteCell[] = [];
  for (const {cells} of sheet.rows()) {
    for (const cell of cells) {
      if (cell.note !== undefined) {
        notes.push({ref: cell.address, row: cell.row, col: cell.col, text: cell.note});
      }
    }
  }
  return notes;
}

/** The `xl/comments{n}.xml` part: a single anonymous author and one comment per noted cell. */
export function commentsXml(notes: readonly NoteCell[]): string {
  const list = notes
    .map((note) => {
      const preserve = needsSpacePreserve(note.text) ? ' xml:space="preserve"' : '';
      return (
        `<comment ref="${note.ref}" authorId="0">` +
        `<text><r><t${preserve}>${escapeText(note.text)}</t></r></text>` +
        '</comment>'
      );
    })
    .join('');
  return (
    XML_DECLARATION +
    `<comments xmlns="${SPREADSHEETML_NS}">` +
    '<authors><author></author></authors>' +
    `<commentList>${list}</commentList>` +
    '</comments>'
  );
}

// VML namespaces and the one shape type (a text box) every note reuses.
const VML_HEADER =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" ' +
  'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>' +
  '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';

/** The `xl/drawings/vmlDrawing{n}.vml` companion: one hidden text-box shape per noted cell. Anchor
 * coordinates place the box a couple of cells down-and-right of its owner; Excel refines them on open,
 * so the values are a sensible starting geometry rather than a pixel-exact layout. */
export function vmlDrawingXml(notes: readonly NoteCell[]): string {
  const shapes = notes
    .map((note, i) => {
      const row0 = note.row - 1;
      const col0 = note.col - 1;
      const anchor = `${col0 + 1}, 15, ${row0}, 2, ${col0 + 3}, 15, ${row0 + 4}, 4`;
      return (
        `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" ` +
        'style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;' +
        `z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
        '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>' +
        '<v:path o:connecttype="none"/>' +
        '<v:textbox style="mso-direction-alt:auto;mso-fit-shape-to-text:t"><div style="text-align:left"></div></v:textbox>' +
        '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
        `<x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>` +
        `<x:Row>${row0}</x:Row><x:Column>${col0}</x:Column></x:ClientData></v:shape>`
      );
    })
    .join('');
  return `${VML_HEADER}${shapes}</xml>`;
}

/** Parse a `comments{n}.xml` part into a map of A1 reference → note text. Text runs within one
 * comment are concatenated; an author-name run is Excel's own convention and is not stripped, so a
 * note reads back as exactly the text that was written. */
export function parseComments(xml: string): Map<string, string> {
  const notes = new Map<string, string>();
  let currentRef: string | undefined;
  let inText = false;
  let buffer = '';
  parseXml(xml, {
    onOpen(name, attrs) {
      const local = localName(name);
      if (local === 'comment') {
        currentRef = attrs.ref;
        buffer = '';
      } else if (local === 'text') {
        inText = true;
      }
    },
    onText(text) {
      if (inText) buffer += text;
    },
    onClose(name) {
      const local = localName(name);
      if (local === 'text') {
        inText = false;
      } else if (local === 'comment' && currentRef !== undefined) {
        notes.set(currentRef, buffer);
        currentRef = undefined;
      }
    },
  });
  return notes;
}

/** Apply parsed notes onto a sheet's cells, addressing each by its A1 reference. */
export function applyNotes(sheet: Worksheet, notes: Map<string, string>): void {
  for (const [ref, text] of notes) sheet.getCell(ref).note = text;
}
