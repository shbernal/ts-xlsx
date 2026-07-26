// Modern threaded comments in the model — the review-style conversations Excel has written since 2018:
// an anchored discussion of authored messages with timestamps, replies, a resolved state, and
// @mentions. Distinct from a legacy note (`cell.note`), which is a single anonymous annotation; Excel's
// own UI draws the same line, calling these "comments" and those "notes". A cell carries one or the
// other, never both.
//
// The wire form lives in two parts (see `io/xlsx/threaded-comments.ts`): the messages per sheet, and a
// workbook-level identity registry the messages point into. The model mirrors that split — a thread's
// authors and mentioned people are {@link Person} entries resolved through
// {@link Workbook.getPerson}, not names duplicated onto every message.

/**
 * A registered identity a threaded comment can point at — an author, or someone `@mentioned` in a
 * message. One `<person>` of the workbook's `xl/persons/person.xml` registry.
 *
 * A single human legitimately has **several** entries: Excel registers a mentioned identity separately
 * from that person's authoring identity, with the same {@link displayName} and {@link userId} but a
 * different {@link id} and a different {@link providerId}. The id is therefore the only identity —
 * see {@link Workbook.getPerson}.
 */
export interface Person {
  /** Brace-wrapped GUID this identity is referenced by. The only field that identifies it. */
  readonly id: string;
  /** The name a spreadsheet app shows — not unique, and not an identity. */
  readonly displayName: string;
  /** Identity-provider handle, `S::<email>::<tenant-guid>` for an AzureAD account. */
  readonly userId?: string;
  /** The provider that registered this entry — `AD` for a directory account, `PeoplePicker` for an
   * identity interned by being mentioned. */
  readonly providerId?: string;
}

/**
 * An `@mention` inside a message: who was named, and the run of {@link Comment.text} that renders as the
 * mention chip.
 *
 * The offsets are only meaningful against that exact text — shift either and a spreadsheet app
 * highlights the wrong words.
 */
export interface Mention {
  /**
   * The mentioned identity, resolved through the workbook registry. Absent when the file names an id
   * the registry does not hold (a mention left dangling by a foreign generator); {@link personId}
   * still says who was meant.
   */
  readonly person?: Person;
  /** The mentioned {@link Person.id} exactly as written, so a dangling mention stays diagnosable. */
  readonly personId: string;
  /** Excel's own id for this mention, preserved so re-emitting it does not invent a new one. */
  readonly mentionId?: string;
  /** 0-based character offset into {@link Comment.text} where the mention starts. */
  readonly startIndex: number;
  /** Length of the mention in characters, **counting the leading `@`** (`@Grace Hopper` is 13). */
  readonly length: number;
}

/** One message of a {@link CommentThread} — what a single person wrote, once. */
export interface Comment {
  /** Brace-wrapped GUID identifying this message, preserved verbatim from the file. */
  readonly id: string;
  /**
   * Who wrote it, resolved through the workbook registry. Absent when the file recorded no author or
   * named an id the registry does not hold; {@link personId} distinguishes those two cases.
   */
  readonly author?: Person;
  /** The author's {@link Person.id} exactly as written; absent when the file recorded no author. */
  readonly personId?: string;
  /**
   * When it was written, verbatim. Excel writes local wall-clock with fractional seconds and no
   * timezone (`2026-07-24T10:56:41.72`), which is not a round-trippable instant — keeping the string
   * spares the reader from inventing a zone the file never stated.
   */
  readonly date?: string;
  /** The message body as plain text. Mention chips are part of it; {@link mentions} spans it. */
  readonly text: string;
  /** The `@mentions` this message carries, in document order. Empty for a message that names no one. */
  readonly mentions: readonly Mention[];
}

/** A conversation anchored to one cell: what was asked, every reply, and whether it was resolved. */
export interface CommentThread {
  /** A1 reference of the cell the conversation hangs off. */
  readonly ref: string;
  /**
   * Whether the conversation was marked resolved. A property of the *thread*: only the head carries
   * the flag on the wire, so a reply never disagrees with the thread it belongs to.
   */
  readonly resolved: boolean;
  /** The opening message first, then its replies in the order they were written. Never empty. */
  readonly comments: readonly Comment[];
}
