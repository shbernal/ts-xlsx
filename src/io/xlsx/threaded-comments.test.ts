import assert from 'node:assert/strict';
import {test} from 'node:test';

import {parsePersons, parseThreadedComments} from './threaded-comments.ts';

// The part bodies below are verbatim from the Excel-authored corpus fixtures under
// `test/corpus/fixtures/threaded-comment-parts-survive-roundtrip/`, so these tests read the real
// wire shape rather than a plausible reconstruction of it. `resolved-multi-author.xlsx` is the
// interesting one: a resolved thread on B1 whose reply is by a second author, and an open thread on
// B2. Anything hand-written here is labelled as such and covers a shape Excel does not produce.

const PERSONS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" ' +
  'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<person displayName="Grace Hopper" id="{1B2C3D4E-5F60-4A71-8B92-0C1D2E3F4A5B}" ' +
  'userId="S::grace@example.com::00000000-0000-0000-0000-000000000000" providerId="AD"/>' +
  '<person displayName="Ada Lovelace" id="{77EC2EBA-7DAF-4C42-8E0A-9244C8EAE97C}" ' +
  'userId="S::ada@example.com::00000000-0000-0000-0000-000000000000" providerId="AD"/>' +
  '</personList>';

const ADA = '{77EC2EBA-7DAF-4C42-8E0A-9244C8EAE97C}';
const GRACE = '{1B2C3D4E-5F60-4A71-8B92-0C1D2E3F4A5B}';
const B1_HEAD = '{A6473DA8-237C-4E27-ADE9-48D3A7CD15A7}';

const THREADED_COMMENTS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" ' +
  'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  `<threadedComment ref="B1" dT="2026-07-26T10:54:00.01" personId="${ADA}" ` +
  `id="${B1_HEAD}" done="1"><text>Is this gross or net of tax?</text></threadedComment>` +
  `<threadedComment ref="B1" dT="2026-07-26T10:54:00.04" personId="${GRACE}" ` +
  `id="{A758ECCA-12FA-49B1-A091-F8BA683AB83C}" parentId="${B1_HEAD}">` +
  '<text>Gross. Confirmed with finance.</text></threadedComment>' +
  `<threadedComment ref="B2" dT="2026-07-26T10:54:00.04" personId="${ADA}" ` +
  'id="{08B3B787-8F24-49B6-8D78-A5F335591EEA}"><text>Where does this figure come from?</text>' +
  '</threadedComment></ThreadedComments>';

// Values verbatim from `mention-in-thread.xlsx` (markup trimmed to the one message under test), whose
// mention Excel itself re-resolved and re-emitted on save. Excel
// renders `@Grace Hopper` as a mention chip over exactly the span `startIndex="0" length="13"` names, so
// `startIndex` is a 0-based character offset into the message text and `length` includes the leading `@`.
// `<mentions>` is not parsed yet — these tests pin that it cannot corrupt the message that carries it.
const MENTION_PERSON = '{BA397017-DD76-4496-AA75-59ADB199950C}';
const MENTION_MESSAGE =
  '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
  `<threadedComment ref="B2" dT="2026-07-26T10:54:00.04" personId="${ADA}" ` +
  'id="{08B3B787-8F24-49B6-8D78-A5F335591EEA}">' +
  '<text>@Grace Hopper Where does this figure come from?</text>' +
  `<mentions><mention mentionpersonId="${MENTION_PERSON}" ` +
  'mentionId="{3F2C1A9E-5B84-4D67-9C2E-71A0D4E8B531}" startIndex="0" length="13"/></mentions>' +
  '</threadedComment></ThreadedComments>';

// The same file's registry. Excel interns a *mentioned* identity as its own entry — note the third
// `<person>`: same displayName and userId as the second, different id, `providerId="PeoplePicker"`.
const MENTION_PERSONS =
  '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
  `<person displayName="Ada Lovelace" id="${ADA}" ` +
  'userId="S::ada@example.com::00000000-0000-0000-0000-000000000000" providerId="AD"/>' +
  `<person displayName="Grace Hopper" id="${GRACE}" ` +
  'userId="S::grace@example.com::00000000-0000-0000-0000-000000000000" providerId="AD"/>' +
  `<person displayName="Grace Hopper" id="${MENTION_PERSON}" ` +
  'userId="S::grace@example.com::00000000-0000-0000-0000-000000000000" providerId="PeoplePicker"/>' +
  '</personList>';

test('every registered person is parsed with the identity attributes the file carried', () => {
  assert.deepStrictEqual(parsePersons(PERSONS), [
    {
      id: GRACE,
      displayName: 'Grace Hopper',
      userId: 'S::grace@example.com::00000000-0000-0000-0000-000000000000',
      providerId: 'AD',
    },
    {
      id: ADA,
      displayName: 'Ada Lovelace',
      userId: 'S::ada@example.com::00000000-0000-0000-0000-000000000000',
      providerId: 'AD',
    },
  ]);
});

test('the personList root is not mistaken for a person entry', () => {
  assert.strictEqual(parsePersons(PERSONS).length, 2);
  assert.deepStrictEqual(parsePersons('<personList/>'), []);
});

test('a person keeps optional identity attributes absent rather than empty', () => {
  const [person] = parsePersons('<personList><person displayName="Anon" id="{A}"/></personList>');
  assert.deepStrictEqual(person, {id: '{A}', displayName: 'Anon'});
});

test('two registry entries for the same human are both kept, since only the id identifies one', () => {
  // Excel registers a mentioned identity separately from the same person's authoring entry, so the
  // registry legitimately holds duplicates by name and userId. Deduplicating on either would drop an
  // entry a mention points at; the id is the only key.
  const persons = parsePersons(MENTION_PERSONS);
  const graces = persons.filter((p) => p.displayName === 'Grace Hopper');
  assert.strictEqual(graces.length, 2, 'both Grace entries survive');
  assert.deepStrictEqual(
    graces.map((p) => [p.id, p.providerId]),
    [
      [GRACE, 'AD'],
      [MENTION_PERSON, 'PeoplePicker'],
    ],
    'they differ only by id and the provider that registered them',
  );
});

test('a person with no id is skipped, since no message could reference it', () => {
  assert.deepStrictEqual(
    parsePersons('<personList><person displayName="Nobody"/></personList>'),
    [],
  );
});

test('a person with an id but no displayName is kept — the id is what messages point at', () => {
  assert.deepStrictEqual(parsePersons('<personList><person id="{A}"/></personList>'), [
    {id: '{A}', displayName: ''},
  ]);
});

test('every message is parsed in document order with its anchor, author and timestamp', () => {
  const messages = parseThreadedComments(THREADED_COMMENTS);
  assert.deepStrictEqual(
    messages.map((m) => [m.ref, m.personId, m.date, m.text]),
    [
      ['B1', ADA, '2026-07-26T10:54:00.01', 'Is this gross or net of tax?'],
      ['B1', GRACE, '2026-07-26T10:54:00.04', 'Gross. Confirmed with finance.'],
      ['B2', ADA, '2026-07-26T10:54:00.04', 'Where does this figure come from?'],
    ],
  );
});

test('the local-time dT is kept verbatim rather than reinterpreted as an instant', () => {
  const [head] = parseThreadedComments(THREADED_COMMENTS);
  assert.strictEqual(head?.date, '2026-07-26T10:54:00.01');
});

test('a reply points at its thread head and the head points at nothing', () => {
  const [head, reply, other] = parseThreadedComments(THREADED_COMMENTS);
  assert.strictEqual(head?.id, B1_HEAD);
  assert.strictEqual(head?.parentId, undefined);
  assert.strictEqual(reply?.parentId, B1_HEAD);
  assert.strictEqual(other?.parentId, undefined);
});

test('done is read off the head alone: a reply and an open thread both report false', () => {
  const [head, reply, openHead] = parseThreadedComments(THREADED_COMMENTS);
  assert.strictEqual(head?.done, true);
  assert.strictEqual(reply?.done, false);
  assert.strictEqual(openHead?.done, false);
});

test('done="true" is honoured alongside done="1"', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}" done="true"><text>x</text>' +
      '</threadedComment><threadedComment ref="A2" id="{B}" done="0"><text>y</text>' +
      '</threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.done),
    [true, false],
  );
});

test('each message gets its own text, with no bleed from the message before it', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>first</text></threadedComment>' +
      '<threadedComment ref="A2" id="{B}"><text>second</text></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.text),
    ['first', 'second'],
  );
});

test('message text is entity-decoded and keeps its line breaks', () => {
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}">' +
      '<text>5 &lt; 6 &amp; "quoted"\nsecond line</text></threadedComment></ThreadedComments>',
  );
  assert.strictEqual(message?.text, '5 < 6 & "quoted"\nsecond line');
});

test('a message that @mentions someone keeps its text intact around the mention', () => {
  const [message] = parseThreadedComments(MENTION_MESSAGE);
  assert.strictEqual(message?.text, '@Grace Hopper Where does this figure come from?');
  assert.strictEqual(message.personId, ADA, 'the author is unaffected by the mention it contains');
});

test('parts that prefix the extension namespace instead of defaulting it still parse', () => {
  // Hand-written: Excel defaults the namespace, but a foreign generator may bind it to a prefix.
  const messages = parseThreadedComments(
    '<tc:ThreadedComments xmlns:tc="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
      '<tc:threadedComment ref="A1" id="{A}" done="1"><tc:text>prefixed</tc:text>' +
      '</tc:threadedComment></tc:ThreadedComments>',
  );
  assert.deepStrictEqual(messages, [{ref: 'A1', id: '{A}', text: 'prefixed', done: true}]);
});

test('a message with no text element parses as empty rather than being dropped', () => {
  // Hand-written: Excel always writes `<text>`, so this is foreign input the reader must tolerate.
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"/>' +
      '<threadedComment ref="A2" id="{B}"></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => [m.ref, m.text]),
    [
      ['A1', ''],
      ['A2', ''],
    ],
  );
});

test('a message that cannot be anchored or replied to is skipped, not guessed at', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment id="{A}"><text>no ref</text></threadedComment>' +
      '<threadedComment ref="A2"><text>no id</text></threadedComment>' +
      '<threadedComment ref="A3" id="{C}"><text>kept</text></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.ref),
    ['A3'],
  );
});

test('a message keeps an absent author and timestamp absent', () => {
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>anon</text>' +
      '</threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(message, {ref: 'A1', id: '{A}', text: 'anon', done: false});
});

test('an empty part parses as no messages and no authors', () => {
  assert.deepStrictEqual(parseThreadedComments('<ThreadedComments/>'), []);
  assert.deepStrictEqual(parsePersons('<personList></personList>'), []);
});
