import assert from 'node:assert/strict';
import {test} from 'node:test';

import {RunAccumulator} from './rich-runs.ts';

// One invariant carries this class, and it is invisible from any single caller: the accumulator is
// emptied when a container opens, so every reader that opens one has to say so. These pin it directly,
// rather than only through the two readers that happen to obey it today.

test('runs bleed into the next string when no container is opened — why the invariant exists', () => {
  const runs = new RunAccumulator();
  runs.beginContainer();
  runs.beginRun();
  runs.appendText('first');
  runs.endRun();

  // No beginContainer here: this is the misuse the method exists to prevent, pinned so a reader that
  // forgets it fails a test rather than shipping a string with a stranger's words in front of it.
  runs.beginRun();
  runs.appendText('second');
  runs.endRun();
  assert.deepStrictEqual(
    runs.runs.map((run) => run.text),
    ['first', 'second'],
  );
});

test('beginContainer discards the previous container runs', () => {
  const runs = new RunAccumulator();
  runs.beginContainer();
  runs.beginRun();
  runs.appendText('first');
  runs.endRun();

  runs.beginContainer();
  runs.beginRun();
  runs.appendText('second');
  runs.endRun();
  assert.deepStrictEqual(
    runs.runs.map((run) => run.text),
    ['second'],
  );
});

test('a value built from one container keeps its own array across the next', () => {
  // Why beginContainer installs a new array instead of emptying the one it has: the consumer reads the
  // runs after the container closes, and a shared array would be cleared out from under it.
  const runs = new RunAccumulator();
  runs.beginContainer();
  runs.beginRun();
  runs.appendText('kept');
  runs.endRun();
  const captured = runs.runs;

  runs.beginContainer();
  assert.deepStrictEqual(
    captured.map((run) => run.text),
    ['kept'],
  );
});

test('beginContainer closes an unterminated run rather than committing it to the next container', () => {
  // A truncated `<is>` leaves a run open. The next container must not inherit it, or the first cell of
  // a damaged sheet bleeds into the second.
  const runs = new RunAccumulator();
  runs.beginContainer();
  runs.beginRun();
  runs.appendText('orphan');

  runs.beginContainer();
  assert.strictEqual(runs.appendText('loose'), false, 'no run is open in the fresh container');
  assert.deepStrictEqual(runs.runs, []);
});

test('a run keeps only the font facets its own rPr set', () => {
  const runs = new RunAccumulator();
  runs.beginContainer();
  runs.beginRun();
  runs.beginProperties();
  runs.applyProperty('b', {});
  runs.appendText('bold');
  runs.endRun();
  // The next run declares no rPr at all, so it must carry no font rather than the previous run's.
  runs.beginRun();
  runs.appendText('plain');
  runs.endRun();
  assert.deepStrictEqual(runs.runs, [{text: 'bold', font: {bold: true}}, {text: 'plain'}]);
});

test('an rPr that sets no facet leaves the run without a font', () => {
  const runs = new RunAccumulator();
  runs.beginContainer();
  runs.beginRun();
  runs.beginProperties();
  runs.endRun();
  assert.deepStrictEqual(runs.runs, [{text: ''}]);
});

test('an rPr child outside a run is ignored rather than fabricating one', () => {
  const runs = new RunAccumulator();
  runs.beginContainer();
  runs.applyProperty('b', {});
  runs.endRun();
  assert.deepStrictEqual(runs.runs, []);
});
