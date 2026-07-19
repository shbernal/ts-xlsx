import assert from 'node:assert/strict';
import {test} from 'node:test';

import {parseStyleTable} from './read-styles.ts';

// The style reader narrows every enumerated attribute through a guard: a valid token passes verbatim,
// an unrecognised one is dropped rather than trusted into the model as a bogus union member.

test('a valid border-edge style is kept; an unrecognised one is dropped', () => {
  const good = parseStyleTable(
    '<styleSheet>' +
      '<borders count="2"><border/><border><left style="thin"/></border></borders>' +
      '<cellXfs count="1"><xf borderId="1"/></cellXfs>' +
      '</styleSheet>',
  );
  assert.deepEqual(good.cellXfs[0]?.border, {left: {style: 'thin'}});

  const bad = parseStyleTable(
    '<styleSheet>' +
      '<borders count="2"><border/><border><left style="frobnicate"/></border></borders>' +
      '<cellXfs count="1"><xf borderId="1"/></cellXfs>' +
      '</styleSheet>',
  );
  assert.equal(bad.cellXfs[0]?.border, undefined, 'the invalid edge is dropped, leaving no border');
});

test('a named underline is kept, "none" reads false, and an unknown token stays a plain underline', () => {
  const read = (val: string): unknown => {
    const table = parseStyleTable(
      `<styleSheet><fonts count="1"><font><u val="${val}"/></font></fonts>` +
        '<cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>',
    );
    return table.cellXfs[0]?.font?.underline;
  };
  assert.equal(read('double'), 'double');
  assert.equal(read('none'), false);
  assert.equal(
    read('squiggly'),
    true,
    'an unrecognised style keeps the underline, drops the token',
  );
});

test('an unrecognised vertAlign or scheme token is dropped from the font', () => {
  const table = parseStyleTable(
    '<styleSheet><fonts count="1">' +
      '<font><b/><vertAlign val="diagonal"/><scheme val="fancy"/></font>' +
      '</fonts><cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>',
  );
  assert.deepEqual(table.cellXfs[0]?.font, {bold: true}, 'only the valid facet survives');
});

test('a valid vertAlign and scheme pass through verbatim', () => {
  const table = parseStyleTable(
    '<styleSheet><fonts count="1">' +
      '<font><vertAlign val="superscript"/><scheme val="minor"/></font>' +
      '</fonts><cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>',
  );
  assert.equal(table.cellXfs[0]?.font?.vertAlign, 'superscript');
  assert.equal(table.cellXfs[0]?.font?.scheme, 'minor');
});

test('an unrecognised alignment token is dropped; a valid one is kept', () => {
  const bad = parseStyleTable(
    '<styleSheet><cellXfs count="1">' +
      '<xf><alignment horizontal="sideways" vertical="floating"/></xf>' +
      '</cellXfs></styleSheet>',
  );
  assert.equal(
    bad.cellXfs[0]?.alignment,
    undefined,
    'both bogus tokens drop, leaving no alignment',
  );

  const good = parseStyleTable(
    '<styleSheet><cellXfs count="1">' +
      '<xf><alignment horizontal="center" vertical="top"/></xf>' +
      '</cellXfs></styleSheet>',
  );
  assert.deepEqual(good.cellXfs[0]?.alignment, {horizontal: 'center', vertical: 'top'});
});
