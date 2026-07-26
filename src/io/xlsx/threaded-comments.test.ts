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

test('a mentions sibling does not leak into the message text', () => {
  // Hand-written: no fixture carries a mention yet, so `<mentions>` is unparsed — but it must not
  // corrupt the text of the message that holds it.
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>@Ada please check</text>' +
      '<mentions><mention mentionpersonId="{P}" mentionId="{M}" startIndex="0" length="4"/>' +
      '</mentions></threadedComment></ThreadedComments>',
  );
  assert.strictEqual(message?.text, '@Ada please check');
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
