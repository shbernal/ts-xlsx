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
// `<mentions>` is deliberately not parsed yet: no fixture carries one, so its attribute names are
// unverified, and guessing wire names is how a reader ends up silently dropping real content.

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
 * each thread's replies following its head. Grouping into threads is the caller's job; this stays
 * faithful to the part. A message without a `ref` or `id` cannot be anchored or replied to and is
 * skipped.
 */
export function parseThreadedComments(xml: string): ParsedThreadedComment[] {
  const messages: ParsedThreadedComment[] = [];
  let open: XmlAttributes | undefined;
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
        } else if (local === 'text' && open !== undefined) {
          inText = true;
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
            const message = threadedCommentFrom(open, text);
            if (message !== undefined) messages.push(message);
          }
          open = undefined;
          text = '';
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
): ParsedThreadedComment | undefined {
  const {ref, id, personId, dT, parentId} = attrs;
  if (ref === undefined || id === undefined) return undefined;
  return {
    ref,
    id,
    text,
    done: boolStrict(attrs.done),
    ...(personId !== undefined ? {personId} : {}),
    ...(dT !== undefined ? {date: dT} : {}),
    ...(parentId !== undefined ? {parentId} : {}),
  };
}
