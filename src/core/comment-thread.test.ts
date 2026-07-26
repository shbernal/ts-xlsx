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

const threadAt = (ref: string): CommentThread => ({
  ref,
  resolved: false,
  comments: [{id: `{HEAD-${ref}}`, text: `about ${ref}`, mentions: []}],
});

test('the person registry keeps both entries of one human, since only the id identifies one', () => {
  const wb = new Workbook();
  wb.restorePersons([GRACE_AUTHOR, GRACE_MENTIONED]);
  assert.equal(wb.persons.length, 2);
  assert.equal(wb.getPerson(GRACE_AUTHOR.id)?.providerId, 'AD');
  assert.equal(wb.getPerson(GRACE_MENTIONED.id)?.providerId, 'PeoplePicker');
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
