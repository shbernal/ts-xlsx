// Modern threaded comments — the review-style conversations Excel has written since 2018, read from and
// written to their two parts: `xl/threadedComments/threadedComment{n}.xml` (per sheet) and
// `xl/persons/person.xml` (per workbook, the author registry).
//
// These are a Microsoft extension, not base ECMA-376, and they are a *separate* feature from legacy
// notes (`comments{n}.xml`, see `comments.ts`) rather than a newer spelling of them. A cell carries
// one or the other, never both — Excel refuses to add a note to a threaded cell and vice versa.
//
// A message identifies its author by `personId` into the registry and its thread by `id`/`parentId`:
// the first message of a thread has no `parentId`, every reply carries the head's `id`. Only the head
// carries `done`, so a thread's resolved state is its head's — a reply never says.
//
// These parsers describe a file that already exists, so they read leniently: an unrecognised or
// missing attribute yields a sensible default rather than a throw, and a message too incomplete to
// anchor is skipped instead of crashing the read. Optional wire attributes stay optional in the
// parsed shape rather than collapsing to `''`, so "the file did not say" never masquerades as a value
// the file contained.
//
// {@link buildCommentThreads} turns the flat message list into the model's threads, resolving each
// author and mention against the workbook's person registry. That grouping is deliberately not the
// parsers' job: they stay faithful to the part, one function per wire form.
//
// The writers are the exact inverse and carry no clock and no id generator: every guid and every
// timestamp is written from the model, verbatim, so the same workbook always serialises to the same
// bytes. A conversation Excel wrote round-trips as itself; one authored in the model carries whatever
// ids and dates the caller supplied.

import {type CellAddress, decodeAddress} from '../../core/address.ts';
import type {Comment, CommentThread, Mention, Person} from '../../core/comment-thread.ts';
import {THREADED_COMMENTS_NS} from './namespaces.ts';
import {escapeAttr, escapeText, XML_DECLARATION} from './xml.ts';
import {boolStrict, localName, parseXml, type XmlAttributes} from './xml-read.ts';

/** A registered author of threaded comments — one `<person>` of `xl/persons/person.xml`. */
export interface ParsedPerson {
  /** Brace-wrapped GUID a message's `personId` points at. */
  readonly id: string;
  readonly displayName: string;
  /** Identity-provider handle, `S::<email>::<tenant-guid>` for an AzureAD account. */
  readonly userId?: string;
  /** Identity provider, e.g. `AD`. */
  readonly providerId?: string;
}

/**
 * One `<mention>` of a message's `<mentions>` block: who was named, and the run of the message text
 * that renders as the mention chip.
 *
 * All four wire attributes are required (Excel rejects a file missing any), and note the lowercase `p`
 * in `mentionpersonId` — the capitalised spelling is not a declared attribute.
 */
export interface ParsedMention {
  /** The mentioned {@link ParsedPerson.id}, from `mentionpersonId`. */
  readonly personId: string;
  /** Excel's own id for the mention itself. Absent only in a file that omitted it. */
  readonly mentionId?: string;
  /**
   * 0-based character offset into the message text. Verified against desktop Excel by rendering: the
   * chip covers exactly `[startIndex, startIndex + length)` of the text.
   */
  readonly startIndex: number;
  /** The mention's length in characters, **including the leading `@`** (`@Grace Hopper` is 13). */
  readonly length: number;
}

/** One message of a threaded conversation — a `<threadedComment>` of a `threadedComment{n}.xml`. */
export interface ParsedThreadedComment {
  /** A1 reference of the cell the whole thread anchors to; every message of a thread repeats it. */
  readonly ref: string;
  /** Brace-wrapped GUID identifying this message, and the `parentId` its replies carry. */
  readonly id: string;
  /** The author's {@link ParsedPerson.id}. Absent in a file that recorded no author. */
  readonly personId?: string;
  /**
   * The `dT` timestamp verbatim. Excel writes local wall-clock with fractional seconds and no
   * timezone (`2026-07-24T10:56:41.72`), which is not a round-trippable instant — keeping the
   * string spares the reader from inventing a zone the file never stated.
   */
  readonly date?: string;
  readonly text: string;
  /** The thread head's {@link ParsedThreadedComment.id}; absent iff this message *is* the head. */
  readonly parentId?: string;
  /**
   * The `done` flag exactly as written. Excel puts it on the head alone and omits it entirely on an
   * open thread (never `done="0"`), so read a thread's resolved state off its head — a reply's
   * `false` here means "did not say", not "not resolved".
   */
  readonly done: boolean;
  /** The `<mentions>` this message carries, in document order; empty when it names no one. */
  readonly mentions: readonly ParsedMention[];
}

/**
 * Parse `xl/persons/person.xml` into its registered authors, in document order. Order carries no
 * meaning — Excel re-sorts the list by person id when it saves — so nothing may depend on it. An
 * entry without an `id` is skipped: no message could reference it.
 */
export function parsePersons(xml: string): ParsedPerson[] {
  const persons: ParsedPerson[] = [];
  parseXml(xml, {
    onOpen(name, attrs) {
      // `personList` shares the prefix, so match the exact local name rather than a `startsWith`.
      if (localName(name) !== 'person') return;
      const {id, displayName, userId, providerId} = attrs;
      if (id === undefined) return;
      persons.push({
        id,
        displayName: displayName ?? '',
        ...(userId !== undefined ? {userId} : {}),
        ...(providerId !== undefined ? {providerId} : {}),
      });
    },
  });
  return persons;
}

// Committing a message from its close handling alone would silently drop `<threadedComment …/>`,
// which fires no close of its own; presenting the self-closing form as an empty element gives it one.
const THREADED_COMMENT_EMPTY_CLOSES: ReadonlySet<string> = new Set(['threadedComment']);

/**
 * Parse a `threadedComment{n}.xml` part into its messages, in document order — thread order, with
 * each thread's replies following its head. Grouping into threads is {@link buildCommentThreads}'s
 * job; this stays faithful to the part. A message without a `ref` or `id` cannot be anchored or
 * replied to and is skipped.
 */
export function parseThreadedComments(xml: string): ParsedThreadedComment[] {
  const messages: ParsedThreadedComment[] = [];
  let open: XmlAttributes | undefined;
  let mentions: ParsedMention[] = [];
  let inText = false;
  let text = '';
  parseXml(
    xml,
    {
      onOpen(name, attrs) {
        const local = localName(name);
        if (local === 'threadedComment') {
          open = attrs;
          text = '';
          mentions = [];
        } else if (local === 'text' && open !== undefined) {
          inText = true;
        } else if (local === 'mention' && open !== undefined) {
          const mention = mentionFrom(attrs);
          if (mention !== undefined) mentions.push(mention);
        }
      },
      onText(chunk) {
        if (inText) text += chunk;
      },
      onClose(name) {
        const local = localName(name);
        if (local === 'text') {
          inText = false;
        } else if (local === 'threadedComment') {
          if (open !== undefined) {
            const message = threadedCommentFrom(open, text, mentions);
            if (message !== undefined) messages.push(message);
          }
          open = undefined;
          text = '';
          mentions = [];
        }
      },
    },
    {closeEmptyElements: THREADED_COMMENT_EMPTY_CLOSES},
  );
  return messages;
}

function threadedCommentFrom(
  attrs: XmlAttributes,
  text: string,
  mentions: readonly ParsedMention[],
): ParsedThreadedComment | undefined {
  const {ref, id, personId, dT, parentId} = attrs;
  if (ref === undefined || id === undefined) return undefined;
  return {
    ref,
    id,
    text,
    done: boolStrict(attrs.done),
    mentions: [...mentions],
    ...(personId !== undefined ? {personId} : {}),
    ...(dT !== undefined ? {date: dT} : {}),
    ...(parentId !== undefined ? {parentId} : {}),
  };
}

// A mention without a target person or a usable span cannot be resolved or rendered, so it is dropped
// rather than carried as a mention over nothing — a `length` of 0 would be an invisible chip, and a
// negative or non-numeric offset would place it outside the text it is supposed to cover.
function mentionFrom(attrs: XmlAttributes): ParsedMention | undefined {
  const personId = attrs.mentionpersonId;
  const startIndex = integerAttribute(attrs.startIndex);
  const length = integerAttribute(attrs.length);
  if (personId === undefined || startIndex === undefined || length === undefined) return undefined;
  // A negative offset points outside the text and a zero length spans nothing, so neither could render.
  if (startIndex < 0 || length <= 0) return undefined;
  return {
    personId,
    startIndex,
    length,
    ...(attrs.mentionId !== undefined ? {mentionId: attrs.mentionId} : {}),
  };
}

// An integer attribute, or undefined when the file did not write a usable one. Blank is rejected before
// `Number` sees it, since `Number('')` is 0 — an empty attribute must not read as offset zero.
function integerAttribute(raw: string | undefined): number | undefined {
  if (raw === undefined || raw.trim() === '') return undefined;
  const value = Number(raw);
  return Number.isInteger(value) ? value : undefined;
}

/**
 * Group a part's messages into the model's threads, resolving every author and mention against the
 * workbook's person registry (`personById`, typically `Workbook.getPerson`).
 *
 * Document order is thread order with each thread's replies following its head, so one pass suffices: a
 * message with no `parentId` opens a thread, and a reply joins the thread its `parentId` names. A reply
 * whose parent is unknown — a dangling `parentId` no Excel file produces — opens a thread of its own
 * rather than being dropped, so a foreign generator's damage costs structure, never content.
 */
export function buildCommentThreads(
  messages: readonly ParsedThreadedComment[],
  personById: (id: string) => Person | undefined,
): CommentThread[] {
  const threads: CommentThread[] = [];
  // The head's own message list, so a reply appends to the thread already published in `threads`.
  const commentsByHeadId = new Map<string, Comment[]>();
  for (const message of messages) {
    const comment = commentFrom(message, personById);
    const siblings =
      message.parentId === undefined ? undefined : commentsByHeadId.get(message.parentId);
    if (siblings !== undefined) {
      siblings.push(comment);
      continue;
    }
    const ref = anchorRef(message.ref);
    if (ref === undefined) continue;
    const comments = [comment];
    commentsByHeadId.set(message.id, comments);
    // Resolved is the head's flag: a reply never carries `done`, so it never contradicts its thread.
    threads.push({ref, resolved: message.done, comments});
  }
  return threads;
}

// The canonical A1 form of a thread's anchor, or undefined when the file wrote something that cannot
// anchor one — a range, a bare row or column, or outright garbage. Canonicalising here is what lets
// every later consumer compare anchors as plain strings (`$B$2` and `B2` are one cell) and keeps a
// foreign file's malformed reference out of the writer, which anchors the thread's legacy fallback by it.
function anchorRef(reference: string): string | undefined {
  let decoded: CellAddress;
  try {
    decoded = decodeAddress(reference);
  } catch {
    return undefined;
  }
  return decoded.col === undefined || decoded.row === undefined ? undefined : decoded.address;
}

function commentFrom(
  message: ParsedThreadedComment,
  personById: (id: string) => Person | undefined,
): Comment {
  const author = message.personId === undefined ? undefined : personById(message.personId);
  return {
    id: message.id,
    text: message.text,
    mentions: message.mentions.map((mention) => mentionOf(mention, personById)),
    ...(author !== undefined ? {author} : {}),
    ...(message.personId !== undefined ? {personId: message.personId} : {}),
    ...(message.date !== undefined ? {date: message.date} : {}),
  };
}

function mentionOf(
  mention: ParsedMention,
  personById: (id: string) => Person | undefined,
): Mention {
  const person = personById(mention.personId);
  return {
    personId: mention.personId,
    startIndex: mention.startIndex,
    length: mention.length,
    ...(person !== undefined ? {person} : {}),
    ...(mention.mentionId !== undefined ? {mentionId: mention.mentionId} : {}),
  };
}

/**
 * Serialise one sheet's conversations into its `xl/threadedComments/threadedComment{n}.xml` part.
 *
 * Messages are written flat, in thread order with each thread's replies after its head — the shape
 * {@link parseThreadedComments} reads back. The head/reply distinction the model holds as array position
 * becomes `parentId` on every reply but the head, and `done="1"` goes on the head alone: only the head
 * carries the flag on the wire, so a reply can never contradict the thread it belongs to. An open thread
 * omits `done` entirely rather than writing `done="0"`, exactly as Excel does.
 *
 * A thread with no messages writes nothing: it has neither text to say nor a head id for its replies and
 * its legacy fallback to hang off.
 */
export function threadedCommentsXml(threads: readonly CommentThread[]): string {
  const messages = threads.flatMap((thread) => {
    const [head, ...replies] = thread.comments;
    if (head === undefined) return [];
    return [
      threadedCommentXml(thread.ref, head, thread.resolved ? ' done="1"' : ''),
      ...replies.map((reply) =>
        threadedCommentXml(thread.ref, reply, ` parentId="${escapeAttr(head.id)}"`),
      ),
    ];
  });
  return (
    XML_DECLARATION +
    `<ThreadedComments xmlns="${THREADED_COMMENTS_NS}">${messages.join('')}</ThreadedComments>`
  );
}

// One `<threadedComment>`. `tail` is the attribute that distinguishes the message's role — `done` for a
// resolved head, `parentId` for a reply, nothing for an open head. A `dT` or `personId` the model never
// held is omitted rather than written empty, so "the file did not say" stays distinguishable from "the
// file said nothing". Every value is escaped: an authored message's text and a foreign file's ids alike
// are untrusted, and an unescaped `"` would end the attribute and reshape the part.
function threadedCommentXml(ref: string, comment: Comment, tail: string): string {
  const date = comment.date === undefined ? '' : ` dT="${escapeAttr(comment.date)}"`;
  const person =
    comment.personId === undefined ? '' : ` personId="${escapeAttr(comment.personId)}"`;
  return (
    `<threadedComment ref="${escapeAttr(ref)}"${date}${person} id="${escapeAttr(comment.id)}"${tail}>` +
    `<text>${escapeText(comment.text)}</text>` +
    mentionsXml(comment.mentions) +
    '</threadedComment>'
  );
}

// The `<mentions>` block, which follows `<text>` in the message. All four `<mention>` attributes are
// required — verified by dropping each in turn and getting `Sch_MissRequiredAttribute` — so a mention the
// model holds without a `mentionId` cannot be written at all. It is dropped rather than given an invented
// id: the `@name` stays in the text and only the chip is lost, whereas an invalid part risks Excel
// repairing the whole conversation away. (Excel always writes the id, so this needs a foreign generator.)
function mentionsXml(mentions: readonly Mention[]): string {
  const entries = mentions.flatMap((mention) =>
    mention.mentionId === undefined
      ? []
      : [
          `<mention mentionpersonId="${escapeAttr(mention.personId)}"` +
            ` mentionId="${escapeAttr(mention.mentionId)}"` +
            ` startIndex="${mention.startIndex}" length="${mention.length}"/>`,
        ],
  );
  return entries.length === 0 ? '' : `<mentions>${entries.join('')}</mentions>`;
}

/**
 * Serialise the workbook's identity registry into `xl/persons/person.xml` — singular and unnumbered,
 * unlike the per-sheet thread parts.
 *
 * Entries are written in registry order, which carries no meaning: Excel re-sorts the list by person id
 * whenever it saves, so this only has to be deterministic, not canonical. `userId`/`providerId` are
 * written when the model holds them; a registry read from a file holds whatever that file stated.
 */
export function personsXml(persons: readonly Person[]): string {
  const entries = persons.map((person) => {
    const userId = person.userId === undefined ? '' : ` userId="${escapeAttr(person.userId)}"`;
    const providerId =
      person.providerId === undefined ? '' : ` providerId="${escapeAttr(person.providerId)}"`;
    return (
      `<person displayName="${escapeAttr(person.displayName)}"` +
      ` id="${escapeAttr(person.id)}"${userId}${providerId}/>`
    );
  });
  return `${XML_DECLARATION}<personList xmlns="${THREADED_COMMENTS_NS}">${entries.join('')}</personList>`;
}
