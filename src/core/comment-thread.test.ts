import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {CommentThread, Person} from './comment-thread.ts';
import {Workbook} from './workbook.ts';

// The two entries Excel writes for one human: an authoring identity, and the separate one it interns
// when that person is @mentioned. Same name, same userId, different id and provider.
const GRACE_AUTHOR: Person = {
  id: '{1B2C3D4E-5F60-4A71-8B92-0C1D2E3F4A5B}',
  displayName: 'Grace Hopper',
  userId: 'S::grace@example.com::00000000-0000-0000-0000-000000000000',
  providerId: 'AD',
};
const GRACE_MENTIONED: Person = {
  id: '{BA397017-DD76-4496-AA75-59ADB199950C}',
  displayName: 'Grace Hopper',
  userId: 'S::grace@example.com::00000000-0000-0000-0000-000000000000',
  providerId: 'PeoplePicker',
};

// A distinct id per thread, in the only spelling the format accepts: brace-wrapped, upper-case, `8-4-4-4-12`
// hex. A readable placeholder like `{HEAD-B2}` is not a legal id and the authoring path rejects it.
let nextThread = 0;
const threadAt = (ref: string): CommentThread => ({
  ref,
  resolved: false,
  comments: [
    {
      id: `{${String(++nextThread).padStart(8, '0')}-0000-4000-8000-000000000000}`,
      text: `about ${ref}`,
      mentions: [],
    },
  ],
});

test('the person registry keeps both entries of one human, since only the id identifies one', () => {
  const wb = new Workbook();
  wb.restorePersons([GRACE_AUTHOR, GRACE_MENTIONED]);
  assert.equal(wb.persons.length, 2);
  assert.equal(wb.getPerson(GRACE_AUTHOR.id)?.providerId, 'AD');
  assert.equal(wb.getPerson(GRACE_MENTIONED.id)?.providerId, 'PeoplePicker');
});

test('registering the same id twice replaces the entry, since the id is the identity', () => {
  const wb = new Workbook();
  wb.addPerson(GRACE_AUTHOR);
  wb.addPerson({...GRACE_AUTHOR, displayName: 'Rear Admiral Grace Hopper'});
  assert.equal(wb.persons.length, 1);
  assert.equal(wb.getPerson(GRACE_AUTHOR.id)?.displayName, 'Rear Admiral Grace Hopper');
});

test('registering one human under two ids keeps both, which is what Excel itself writes', () => {
  const wb = new Workbook();
  wb.addPerson(GRACE_AUTHOR);
  wb.addPerson(GRACE_MENTIONED);
  assert.deepEqual(
    wb.persons.map((person) => person.providerId),
    ['AD', 'PeoplePicker'],
  );
});

test('an id the registry does not hold resolves to nothing rather than a near match', () => {
  const wb = new Workbook();
  wb.restorePersons([GRACE_AUTHOR]);
  assert.equal(wb.getPerson('{NOT-REGISTERED}'), undefined);
});

test('a workbook with no threaded comments has an empty person registry', () => {
  assert.deepEqual(new Workbook().persons, []);
});

test('restoring the registry replaces it, never accumulating a stale identity', () => {
  const wb = new Workbook();
  wb.restorePersons([GRACE_AUTHOR]);
  wb.restorePersons([GRACE_MENTIONED]);
  assert.deepEqual(
    wb.persons.map((person) => person.id),
    [GRACE_MENTIONED.id],
  );
  assert.equal(wb.getPerson(GRACE_AUTHOR.id), undefined);
});

test('a sheet exposes its threads in the order they were read', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.restoreCommentThreads([threadAt('B1'), threadAt('B2')]);
  assert.deepEqual(
    sheet.commentThreads.map((thread) => thread.ref),
    ['B1', 'B2'],
  );
});

test('a sheet with no threaded comments has no threads', () => {
  assert.deepEqual(new Workbook().addWorksheet('S').commentThreads, []);
});

test('a thread is found by its anchor, absolute reference or not', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.restoreCommentThreads([threadAt('B2')]);
  assert.equal(sheet.commentThreadAt('B2')?.ref, 'B2');
  assert.equal(sheet.commentThreadAt('$B$2')?.ref, 'B2');
  assert.equal(sheet.commentThreadAt('C3'), undefined);
});

test('looking for a thread does not materialise the cell it asks about', () => {
  // A read-only query that grew the grid would change what the writer emits — a phantom empty cell
  // appearing merely because something inspected the sheet.
  const sheet = new Workbook().addWorksheet('S');
  sheet.commentThreadAt('B2');
  assert.equal(sheet.hasCell(2, 2), false);
});

test('commentThreadAt rejects a reference that is not a single cell', () => {
  const sheet = new Workbook().addWorksheet('S');
  assert.throws(() => sheet.commentThreadAt('A'), /not a single-cell reference/);
  assert.throws(() => sheet.commentThreadAt('2'), /not a single-cell reference/);
});

test('an authored thread keeps the conversation but canonicalises its anchor', () => {
  // The anchor is compared as a plain string by both `commentThreadAt` and the writer's fallback comment,
  // so `$B$2` and `B2` must not become two anchors for one cell.
  const sheet = new Workbook().addWorksheet('S');
  sheet.addCommentThread({...threadAt('$B$2'), resolved: true});
  const [thread] = sheet.commentThreads;
  assert.equal(thread?.ref, 'B2');
  assert.equal(thread?.resolved, true);
  assert.deepEqual(
    thread?.comments.map((comment) => comment.text),
    ['about $B$2'],
  );
  assert.equal(sheet.commentThreadAt('B2')?.ref, 'B2');
});

test('authoring accumulates threads rather than replacing them, unlike reader restoration', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addCommentThread(threadAt('B1'));
  sheet.addCommentThread(threadAt('C3'));
  assert.deepEqual(
    sheet.commentThreads.map((thread) => thread.ref),
    ['B1', 'C3'],
  );
});

test('an authored id is normalised to the one spelling the format accepts', () => {
  // Verified against the OOXML schema: `person/@id`, a message's `id`/`personId`/`parentId` and a mention's
  // `mentionpersonId`/`mentionId` are each pinned to `\{[0-9A-F]{8}-…\}`. A bare GUID is rejected and so is
  // a lower-case one — which is exactly what `crypto.randomUUID()` returns, so the authoring path
  // normalises rather than refusing the one obvious way to make an id in JavaScript.
  const wb = new Workbook();
  wb.addPerson({id: 'aaaaaaaa-1111-2222-3333-444444444444', displayName: 'Ada'});
  assert.deepEqual(
    wb.persons.map((person) => person.id),
    ['{AAAAAAAA-1111-2222-3333-444444444444}'],
  );
  assert.ok(
    wb.getPerson('{AAAAAAAA-1111-2222-3333-444444444444}'),
    'found under the normalised id',
  );

  const sheet = wb.addWorksheet('S');
  sheet.addCommentThread({
    ref: 'A1',
    resolved: false,
    comments: [
      {
        id: 'bbbbbbbb-1111-2222-3333-444444444444',
        personId: 'aaaaaaaa-1111-2222-3333-444444444444',
        text: '@Ada who owns this?',
        mentions: [
          {
            personId: 'aaaaaaaa-1111-2222-3333-444444444444',
            mentionId: '{cccccccc-1111-2222-3333-444444444444}',
            startIndex: 0,
            length: 4,
          },
        ],
      },
    ],
  });
  const [comment] = sheet.commentThreads[0]?.comments ?? [];
  assert.equal(comment?.id, '{BBBBBBBB-1111-2222-3333-444444444444}');
  assert.equal(comment?.personId, '{AAAAAAAA-1111-2222-3333-444444444444}');
  assert.equal(comment?.mentions[0]?.personId, '{AAAAAAAA-1111-2222-3333-444444444444}');
  assert.equal(comment?.mentions[0]?.mentionId, '{CCCCCCCC-1111-2222-3333-444444444444}');
});

test('authoring rejects an id that is not a GUID, rather than writing a file Excel repairs', () => {
  const wb = new Workbook();
  assert.throws(() => wb.addPerson({id: 'ada@example.com', displayName: 'Ada'}), /must be a GUID/);
  const sheet = wb.addWorksheet('S');
  const thread = threadAt('A1');
  assert.throws(
    () => sheet.addCommentThread({...thread, comments: [{...thread.comments[0]!, id: 'thread-1'}]}),
    /a comment id must be a GUID/,
  );
  assert.throws(
    () =>
      sheet.addCommentThread({
        ...thread,
        comments: [{...thread.comments[0]!, personId: 'Ada Lovelace'}],
      }),
    /a comment's author id must be a GUID/,
  );
  assert.deepEqual(sheet.commentThreads, [], 'nothing half-added survives a rejection');
});

test('a thread cannot be anchored to anything but a single cell', () => {
  // A range, a bare column, or a bare row names no one cell to hang the conversation off — and the writer
  // would have nothing to anchor its legacy fallback comment or VML shape to.
  const sheet = new Workbook().addWorksheet('S');
  for (const ref of ['A1:B2', 'A', '2', 'nonsense']) {
    assert.throws(() => sheet.addCommentThread(threadAt(ref)), SyntaxError, ref);
  }
  assert.deepEqual(sheet.commentThreads, []);
});
