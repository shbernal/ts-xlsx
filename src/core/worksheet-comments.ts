// The threaded conversations on a worksheet, lifted off `Worksheet` into the slice they are.
//
// Excel's modern review comments: an author, a timestamp, replies, a resolved state, `@mentions`.
// Distinct from a cell's legacy note, which lives on the cell because that is where the file puts
// it. A thread hangs off exactly one cell and never reaches the grid for anything else, so its only
// edge outside itself is the sheet's name, which one error message quotes and which arrives as an
// accessor rather than a reference to the sheet.
//
// See the delegation rule in `docs/architecture.md`: the public accessors stay on `Worksheet` with
// their doc comments and become one-line calls into this.

import {quoted} from '../errors.ts';
import {decodeCellRef, encodeAddress} from './address.ts';
import {type CommentThread, commentThreadGuid, commentThreadOffset} from './comment-thread.ts';
import {replaceContents} from './containers.ts';
import {isDeletedSpan, shiftIndex} from './grid-shift.ts';

export class WorksheetComments {
  // Quoted by the duplicate-id refusal, which has to say *which* sheet already holds the id.
  readonly #sheetName: () => string;

  readonly #threads: CommentThread[] = [];
  // Every comment id this sheet holds, so `add` tests against a set rather than rebuilding one from
  // every message on the sheet. Rebuilt from scratch by `restore`, which replaces the threads whole;
  // `shift` only re-anchors, and an id does not move with its cell.
  readonly #takenIds = new Set<string>();

  constructor(sheetName: () => string) {
    this.#sheetName = sheetName;
  }

  get threads(): readonly CommentThread[] {
    return this.#threads;
  }

  add(thread: CommentThread): void {
    // The ids this thread claims, kept apart from the sheet's until the whole thread validates: every
    // message is checked before any of it is stored, so a rejection leaves the sheet untouched rather
    // than half-carrying a conversation whose remaining messages were refused.
    const claimed = new Set<string>();
    const comments = thread.comments.map((comment) => {
      const id = commentThreadGuid(comment.id, 'a comment id');
      if (this.#takenIds.has(id) || claimed.has(id)) {
        throw new SyntaxError(
          `a comment id must be unique within a sheet, but ${quoted(id)} is already used on ${quoted(this.#sheetName())}: ` +
            'a reply and the legacy fallback comment both find their thread by it',
        );
      }
      claimed.add(id);
      return {
        ...comment,
        id,
        ...(comment.personId !== undefined
          ? {personId: commentThreadGuid(comment.personId, "a comment's author id")}
          : {}),
        mentions: comment.mentions.map((mention) => ({
          ...mention,
          personId: commentThreadGuid(mention.personId, "a mention's person id"),
          startIndex: commentThreadOffset(mention.startIndex, "a mention's startIndex"),
          length: commentThreadOffset(mention.length, "a mention's length"),
          ...(mention.mentionId !== undefined
            ? {mentionId: commentThreadGuid(mention.mentionId, 'a mention id')}
            : {}),
        })),
      };
    });
    this.#threads.push({...thread, ref: anchorRef(thread.ref), comments});
    for (const id of claimed) this.#takenIds.add(id);
  }

  at(reference: string): CommentThread | undefined {
    const anchor = anchorRef(reference);
    return this.#threads.find((thread) => thread.ref === anchor);
  }

  /**
   * Re-anchor every conversation through a row or column splice. A thread hangs off exactly one cell,
   * so it moves as a point, not a rectangle, and a thread whose cell the splice deleted is dropped
   * along with its messages. Both halves matter: a cell's legacy note is cell state and travels with
   * the cell, so a thread left behind would put a note on one cell and its conversation on another,
   * a pairing the writer emits and Excel refuses.
   */
  shift(axis: 'row' | 'col', start: number, count: number, delta: number): void {
    const survivors: CommentThread[] = [];
    for (const thread of this.#threads) {
      const {col, row} = decodeCellRef(thread.ref);
      const line = axis === 'row' ? row : col;
      if (isDeletedSpan(line, line, start, count)) continue;
      const moved = shiftIndex(line, start, count, delta, axis);
      if (moved === line) {
        survivors.push(thread);
        continue;
      }
      const ref = axis === 'row' ? encodeAddress(col, moved) : encodeAddress(moved, row);
      survivors.push({...thread, ref});
    }
    replaceContents(this.#threads, survivors);
  }

  // The reader's channel: a sheet read from a file arrives with its threads already validated by
  // the codec that parsed them, so they replace rather than re-enter `add`.
  restore(threads: readonly CommentThread[]): void {
    replaceContents(this.#threads, threads);
    this.#takenIds.clear();
    for (const thread of threads) {
      for (const comment of thread.comments) this.#takenIds.add(comment.id);
    }
  }
}

// The canonical A1 form of a conversation's anchor. A thread hangs off one cell, and both the writer's
// fallback comment and {@link WorksheetComments.at} compare anchors as plain strings, so `$B$2` and
// `B2` must not be two anchors.
function anchorRef(reference: string): string {
  const {col, row} = decodeCellRef(reference);
  return encodeAddress(col, row);
}
