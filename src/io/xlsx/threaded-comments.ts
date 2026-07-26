// Modern threaded comments — the review-style conversations Excel has written since 2018, read from
// their two parts: `xl/threadedComments/threadedComment{n}.xml` (per sheet) and
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

import type {Comment, CommentThread, Mention, Person} from '../../core/comment-thread.ts';
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
    const comments = [comment];
    commentsByHeadId.set(message.id, comments);
    // Resolved is the head's flag: a reply never carries `done`, so it never contradicts its thread.
    threads.push({ref: message.ref, resolved: message.done, comments});
  }
  return threads;
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
