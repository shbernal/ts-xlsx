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

import {decodeCellRef, encodeAddress} from './address.ts';
import {type CommentThread, commentThreadGuid, commentThreadOffset} from './comment-thread.ts';
import {replaceContents} from './containers.ts';

export class WorksheetComments {
  // Quoted by the duplicate-id refusal, which has to say *which* sheet already holds the id.
  readonly #sheetName: () => string;

  readonly #threads: CommentThread[] = [];

  constructor(sheetName: () => string) {
    this.#sheetName = sheetName;
  }

  get threads(): readonly CommentThread[] {
    return this.#threads;
  }

  add(thread: CommentThread): void {
    const taken = new Set(
      this.#threads.flatMap((held) => held.comments.map((comment) => comment.id)),
    );
    // Every message is validated before any of it is stored, so a rejection leaves the sheet untouched
    // rather than half-carrying a conversation whose remaining messages were refused.
    const comments = thread.comments.map((comment) => {
      const id = commentThreadGuid(comment.id, 'a comment id');
      if (taken.has(id)) {
        throw new SyntaxError(
          `a comment id must be unique within a sheet, but "${id}" is already used on "${this.#sheetName()}": ` +
            'a reply and the legacy fallback comment both find their thread by it',
        );
      }
      taken.add(id);
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
  }

  at(reference: string): CommentThread | undefined {
    const anchor = anchorRef(reference);
    return this.#threads.find((thread) => thread.ref === anchor);
  }

  // The reader's channel: a sheet read from a file arrives with its threads already validated by
  // the codec that parsed them, so they replace rather than re-enter `add`.
  restore(threads: readonly CommentThread[]): void {
    replaceContents(this.#threads, threads);
  }
}

// The canonical A1 form of a conversation's anchor. A thread hangs off one cell, and both the writer's
// fallback comment and {@link WorksheetComments.at} compare anchors as plain strings, so `$B$2` and
// `B2` must not be two anchors.
function anchorRef(reference: string): string {
  const {col, row} = decodeCellRef(reference);
  return encodeAddress(col, row);
}
