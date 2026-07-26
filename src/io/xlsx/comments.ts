// Cell comments — the `xl/comments{n}.xml` part, its `xl/drawings/vmlDrawing{n}.vml` companion, and the
// reader that maps a comment back onto its cell.
//
// A comment is anchored to a cell by A1 reference and rendered by Excel as a floating box. The box's
// geometry lives in a legacy VML drawing (the pre-DrawingML shape format Excel still requires here);
// the text lives in the comments part. Both are emitted together — a comments part with no matching
// `<legacyDrawing>`/VML reads as text but renders nothing, so we never split them.
//
// Two different things share this one wire form:
//   • a user's **note** (`cell.note`) — a single anonymous annotation, the whole of what the part held
//     before 2018;
//   • the legacy **fallback** Excel writes beside every modern threaded comment (see
//     `threaded-comments.ts`), so a pre-2018 reader still sees the conversation. Its text is a fixed
//     boilerplate wrapping a copy of the thread, and its author is a synthetic `tc={headId}` entry.
//
// That `tc=` author and the comment's `xr:uid` are how Excel binds a cell back to its thread — not
// decoration. Verified against desktop Excel: a package whose threadedComment part, persons registry,
// relationships and content types all survive intact still reads back as ordinary notes with zero
// threads once those two are lost. So the fallback is *derived from the thread model* on write and
// *suppressed on read*, rather than round-tripped as a plain note.

import {decodeAddress} from '../../core/address.ts';
import type {CommentThread} from '../../core/comment-thread.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {MARKUP_COMPATIBILITY_NS, REVISION_NS, SPREADSHEETML_NS} from './namespaces.ts';
import {escapeAttr, escapeText, textElement, XML_DECLARATION} from './xml.ts';
import {localName, parseXml} from './xml-read.ts';

/** A comment bound for `comments{n}.xml`, paired with the coordinates the VML anchor needs. */
export interface CommentCell {
  readonly ref: string;
  /** 1-based row of the commented cell. */
  readonly row: number;
  /** 1-based column of the commented cell. */
  readonly col: number;
  readonly text: string;
  /**
   * The {@link CommentThread} head id this comment is the legacy fallback for, absent for a user's own
   * note. Present means the comment is emitted with a synthetic `tc={id}` author and an `xr:uid`, the
   * pair Excel resolves the cell's thread through.
   */
  readonly threadId?: string;
}

/**
 * Gather every comment a sheet must write: its cells' notes, plus one legacy fallback per conversation
 * in `threads`. A comment anchors to its cell regardless of the cell's value, so a note (or a thread) on
 * an otherwise-empty cell is collected too.
 *
 * `threads` is the conversations the *package* will carry, not simply the ones the sheet holds — the
 * caller decides, because a fallback beside a thread whose `threadedComment` part is missing is worse
 * than no fallback at all: verified against desktop Excel, such a comment shows as neither a thread nor
 * a note, so the text disappears entirely.
 *
 * Ordered by cell, row-major, the way Excel writes the list — so a fallback lands interleaved among the
 * notes rather than appended after them, and the VML shapes follow the same order.
 */
export function collectComments(
  sheet: Worksheet,
  threads: readonly CommentThread[],
): CommentCell[] {
  const fallbacks = threadFallbacks(threads);
  const anchored = new Set(fallbacks.map((fallback) => fallback.ref));
  const comments = [...fallbacks];
  for (const {cells} of sheet.rows()) {
    for (const cell of cells) {
      // Excel refuses to put a note and a thread on one cell, so a file carrying both (only a foreign
      // generator or a hand-edit makes one) is written back as the thread alone: two comments on one ref
      // is a shape Excel repairs by dropping both, which would lose the conversation as well as the note.
      if (cell.note === undefined || anchored.has(cell.address)) continue;
      comments.push({ref: cell.address, row: cell.row, col: cell.col, text: cell.note});
    }
  }
  return comments.sort((a, b) => a.row - b.row || a.col - b.col);
}

// One legacy fallback per conversation, keyed to the thread head whose id binds it. A thread with no
// messages has nothing to write and no id to bind by, so it contributes none.
function threadFallbacks(threads: readonly CommentThread[]): CommentCell[] {
  const fallbacks: CommentCell[] = [];
  for (const thread of threads) {
    const head = thread.comments[0];
    const {col, row} = decodeAddress(thread.ref);
    if (head === undefined || col === undefined || row === undefined) continue;
    fallbacks.push({ref: thread.ref, row, col, text: fallbackText(thread), threadId: head.id});
  }
  return fallbacks;
}

// The boilerplate Excel puts in front of every fallback, captured verbatim from an Excel-authored file.
// It is what a pre-2018 reader shows the user, so it is reproduced exactly rather than paraphrased.
const FALLBACK_PREAMBLE =
  '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any ' +
  'edits to it will get removed if the file is opened in a newer version of Excel. Learn more: ' +
  'https://go.microsoft.com/fwlink/?linkid=870924\n\n';

// A whole conversation flattened into the one comment a pre-2018 reader can render: the opening message
// under `Comment:`, then each reply under its own `Reply:`, every body indented four spaces. Verified
// against desktop Excel for a thread with three replies — `Reply:` repeats per reply rather than the
// replies being joined under one heading.
function fallbackText(thread: CommentThread): string {
  const [head, ...replies] = thread.comments;
  const body = replies.map((reply) => `\nReply:\n    ${reply.text}`).join('');
  return `${FALLBACK_PREAMBLE}Comment:\n    ${head?.text ?? ''}${body}`;
}

// `xr:uid` lives in the 2014 revision namespace, declared `mc:Ignorable` exactly as Excel declares it so
// a consumer that does not know the prefix skips the attribute instead of rejecting the part.
const REVISION_NS_ATTRS = ` xmlns:mc="${MARKUP_COMPATIBILITY_NS}" mc:Ignorable="xr" xmlns:xr="${REVISION_NS}"`;

/**
 * The `xl/comments{n}.xml` part.
 *
 * Authors are laid out the way Excel lays them out: one synthetic `tc={headId}` entry per threaded
 * conversation first, then a single anonymous author shared by every note (the model carries no note
 * author). Each comment points at its own author by index, and a fallback additionally carries the
 * `xr:uid` naming its thread — the pair that keeps Excel treating the cell as threaded.
 */
export function commentsXml(comments: readonly CommentCell[]): string {
  const authorIdByThreadId = new Map<string, number>();
  for (const {threadId} of comments) {
    if (threadId !== undefined && !authorIdByThreadId.has(threadId)) {
      authorIdByThreadId.set(threadId, authorIdByThreadId.size);
    }
  }
  const noteAuthorId = authorIdByThreadId.size;
  const hasNote = comments.some((comment) => comment.threadId === undefined);
  const authors =
    [...authorIdByThreadId.keys()].map((id) => `<author>tc=${escapeText(id)}</author>`).join('') +
    (hasNote ? '<author></author>' : '');

  const list = comments
    .map((comment) => {
      const {threadId} = comment;
      const authorId = threadId === undefined ? noteAuthorId : authorIdByThreadId.get(threadId);
      const uid = threadId === undefined ? '' : ` xr:uid="${escapeAttr(threadId)}"`;
      return (
        `<comment ref="${comment.ref}" authorId="${authorId}"${uid}>` +
        `<text><r>${textElement(comment.text)}</r></text>` +
        '</comment>'
      );
    })
    .join('');

  return (
    XML_DECLARATION +
    `<comments xmlns="${SPREADSHEETML_NS}"${authorIdByThreadId.size > 0 ? REVISION_NS_ATTRS : ''}>` +
    `<authors>${authors}</authors>` +
    `<commentList>${list}</commentList>` +
    '</comments>'
  );
}

// VML namespaces and the one shape type (a text box) every comment reuses.
const VML_HEADER =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" ' +
  'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>' +
  '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';

/** The `xl/drawings/vmlDrawing{n}.vml` companion: one hidden text-box shape per comment, in the same
 * order as the comments part. Anchor coordinates place the box a couple of cells down-and-right of its
 * owner; Excel refines them on open, so the values are a sensible starting geometry rather than a
 * pixel-exact layout. A thread's fallback shape is `ObjectType="Note"` like any other — Excel draws the
 * threaded-comment card itself and only needs the shape to exist. */
export function vmlDrawingXml(comments: readonly CommentCell[]): string {
  const shapes = comments
    .map((comment, i) => {
      const row0 = comment.row - 1;
      const col0 = comment.col - 1;
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

/** One `<comment>` read back from a comments part. */
export interface ParsedComment {
  readonly text: string;
  /**
   * The thread head id this comment is the legacy fallback for, read off its synthetic `tc={headId}`
   * author; absent for a user's own note.
   */
  readonly threadId?: string;
}

// A comment names its author by index into `<authors>`, so an empty entry must still occupy its slot —
// presenting the self-closing `<author/>` an author-less file writes as an empty element gives it the
// close that pushes it. Without this every later index would shift by one and a note could inherit a
// thread's `tc=` author.
const COMMENT_EMPTY_CLOSES: ReadonlySet<string> = new Set(['author']);

// The author string marking a comment as a thread's legacy fallback: `tc={headThreadId}`.
const THREAD_AUTHOR_PREFIX = 'tc=';

/**
 * Parse a `comments{n}.xml` part into a map of A1 reference → comment. Text runs within one comment are
 * concatenated; an author-name run is Excel's own convention and is not stripped, so a note reads back
 * as exactly the text that was written.
 */
export function parseComments(xml: string): Map<string, ParsedComment> {
  const comments = new Map<string, ParsedComment>();
  const authors: string[] = [];
  let currentRef: string | undefined;
  let currentAuthorId: string | undefined;
  let capture: 'author' | 'text' | undefined;
  let buffer = '';
  parseXml(
    xml,
    {
      onOpen(name, attrs) {
        const local = localName(name);
        if (local === 'comment') {
          currentRef = attrs.ref;
          currentAuthorId = attrs.authorId;
          buffer = '';
        } else if (local === 'author') {
          capture = 'author';
          buffer = '';
        } else if (local === 'text') {
          capture = 'text';
        }
      },
      onText(text) {
        if (capture !== undefined) buffer += text;
      },
      onClose(name) {
        const local = localName(name);
        if (local === 'author') {
          authors.push(buffer);
          capture = undefined;
          buffer = '';
        } else if (local === 'text') {
          capture = undefined;
        } else if (local === 'comment' && currentRef !== undefined) {
          const threadId = threadIdOf(authors[Number(currentAuthorId)]);
          comments.set(currentRef, {
            text: buffer,
            ...(threadId !== undefined ? {threadId} : {}),
          });
          currentRef = undefined;
          currentAuthorId = undefined;
        }
      },
    },
    {closeEmptyElements: COMMENT_EMPTY_CLOSES},
  );
  return comments;
}

// A missing or non-numeric `authorId` indexes nothing, so `authors[NaN]` is undefined and the comment
// reads as a plain note — the safe direction, since mistaking a note for a fallback would delete it.
function threadIdOf(author: string | undefined): string | undefined {
  if (author === undefined || !author.startsWith(THREAD_AUTHOR_PREFIX)) return undefined;
  const id = author.slice(THREAD_AUTHOR_PREFIX.length);
  return id === '' ? undefined : id;
}

/**
 * Apply a parsed comments part onto a sheet's cells as notes, addressing each by its A1 reference.
 *
 * A thread's legacy fallback is not a note and does not become one: its text is boilerplate wrapping a
 * copy of the conversation, so surfacing it as `cell.note` hands the caller garbage — and on write it
 * would be re-emitted as a plain note, destroying the `tc=`/`xr:uid` binding and leaving Excel unable to
 * see the thread at all.
 *
 * Suppressed only for a conversation the reader actually holds: a file whose thread part is missing or
 * damaged has nothing else left, so there the boilerplate is kept rather than the content lost. Call
 * after the sheet's threads are restored, since that is what this reads to decide.
 */
export function applyNotes(sheet: Worksheet, comments: ReadonlyMap<string, ParsedComment>): void {
  const headIds = new Set(
    sheet.commentThreads.flatMap((thread) => {
      const head = thread.comments[0];
      return head === undefined ? [] : [head.id];
    }),
  );
  for (const [ref, comment] of comments) {
    if (comment.threadId !== undefined && headIds.has(comment.threadId)) continue;
    sheet.getCell(ref).note = comment.text;
  }
}
