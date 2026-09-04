import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Cell} from './cell.ts';
import {clearCoveredValues, decodeSqrefRects, shiftSqref} from './merge.ts';

// Build the row/column store `clearCoveredValues` walks, from `ref -> value` pairs.
function grid(cells: Readonly<Record<string, string>>): Map<number, Map<number, Cell>> {
  const rows = new Map<number, Map<number, Cell>>();
  for (const [ref, value] of Object.entries(cells)) {
    const [letters = '', digits = ''] = (/^([A-Z]+)(\d+)$/.exec(ref) ?? []).slice(1);
    const col = letters.charCodeAt(0) - 64;
    const row = Number(digits);
    const cell = new Cell(row, col);
    cell.value = value;
    let cols = rows.get(row);
    if (cols === undefined) {
      cols = new Map<number, Cell>();
      rows.set(row, cols);
    }
    cols.set(col, cell);
  }
  return rows;
}

// A `Map` that charges for every entry it hands out and every key it is asked for, and refuses to
// keep counting past a ceiling no correct walk can reach. Timing would measure the machine; this
// measures the algorithm.
class CountedMap<V> extends Map<number, V> {
  readonly #budget: {steps: number};

  constructor(budget: {steps: number}) {
    super();
    this.#budget = budget;
  }

  #charge(): void {
    if (++this.#budget.steps > 10_000_000)
      throw new Error('walked the declared area, not the cells');
  }

  override get(key: number): V | undefined {
    this.#charge();
    return super.get(key);
  }

  override *[Symbol.iterator](): MapIterator<[number, V]> {
    for (const entry of super.entries()) {
      this.#charge();
      yield entry;
    }
  }
}

const valuesOf = (rows: Map<number, Map<number, Cell>>): string[] =>
  [...rows.values()]
    .flatMap((cols) => [...cols.values()])
    .map((cell) => `${cell.address}=${JSON.stringify(cell.value)}`)
    .sort((a, b) => a.localeCompare(b));

test('decodeSqrefRects splits a sqref on whitespace and decodes each area', () => {
  assert.deepEqual(decodeSqrefRects('A1:B2  C3:C4'), [
    {top: 1, left: 1, bottom: 2, right: 2},
    {top: 3, left: 3, bottom: 4, right: 3},
  ]);
});

// A whole column contains a cell anywhere down it, so the missing edge opens rather than clamping to
// the last decoded coordinate: `Infinity` is what makes `at()` a containment test rather than a
// bounded-rectangle test.
test('decodeSqrefRects opens a missing edge to Infinity rather than clamping it', () => {
  assert.deepEqual(decodeSqrefRects('B:B'), [{top: 1, left: 2, bottom: Infinity, right: 2}]);
  assert.deepEqual(decodeSqrefRects('3:3'), [{top: 3, left: 1, bottom: 3, right: Infinity}]);
});

test('a single-cell area decodes to a degenerate rectangle', () => {
  assert.deepEqual(decodeSqrefRects('D7'), [{top: 7, left: 4, bottom: 7, right: 4}]);
});

// The byte-clean round-trip guarantee, and the one thing about `shiftSqref` a refactor could quietly
// break: `B:B` and `B1:B1048576` decode identically, so re-encoding an area the splice did not move
// would rewrite a foreign producer's spelling into ours.
test('shiftSqref returns an unmoved area as its original text, not as a re-encoding', () => {
  assert.equal(shiftSqref('B1:B100', {axis: 'row', start: 900, count: 0, delta: 5}), 'B1:B100');
  assert.equal(
    shiftSqref('$A$1:$C$3', {axis: 'row', start: 90, count: 0, delta: 5}),
    '$A$1:$C$3',
    'anchors survive too',
  );
  assert.equal(
    shiftSqref('A1:C3 F1', {axis: 'col', start: 90, count: 0, delta: 5}),
    'A1:C3 F1',
    'and every area of a list',
  );
});

test('an area unbounded on the spliced axis covers every line of it, so it cannot move', () => {
  // `B:B` after a row insert is still `B:B`, never `B2:B1048577`.
  assert.equal(shiftSqref('B:B', {axis: 'row', start: 1, count: 0, delta: 5}), 'B:B');
  assert.equal(shiftSqref('3:3', {axis: 'col', start: 1, count: 0, delta: 5}), '3:3');
  // The other axis of the same area is bounded, so a splice on *that* axis does move it.
  assert.equal(shiftSqref('B:B', {axis: 'col', start: 1, count: 0, delta: 5}), 'G:G');
});

test('a moved single-cell area stays a single cell', () => {
  // `B5` must not come back as `B6:B6`: the spelling is the caller's, and a range where there was a
  // cell is a different reference to every consumer that compares them as text.
  assert.equal(shiftSqref('B5', {axis: 'row', start: 1, count: 0, delta: 1}), 'B6');
  assert.equal(
    shiftSqref('B5:C5', {axis: 'row', start: 1, count: 0, delta: 1}),
    'B6:C6',
    'a real range still re-encodes as one',
  );
});

test('shiftSqref drops the areas a delete swallowed and keeps the rest', () => {
  // Rows 5..7 deleted: the area inside them is gone, the one below pulls up.
  assert.equal(shiftSqref('A5:A7 A9:A9', {axis: 'row', start: 5, count: 3, delta: -3}), 'A6:A6');
});

test('shiftSqref returns undefined when the splice deleted every area it named', () => {
  // An empty `sqref` is not writable, so the entry holding it must go with it.
  assert.equal(shiftSqref('A5:A7 B6', {axis: 'row', start: 5, count: 3, delta: -3}), undefined);
});

test('an area the machine cannot read is returned untouched rather than dropped', () => {
  // The `sqref` is still the file's own text; moving what cannot be decoded would be a guess, and
  // dropping it would lose a region on the strength of that guess.
  assert.equal(shiftSqref('junk!! A9', {axis: 'row', start: 1, count: 0, delta: 1}), 'junk!! A10');
});

test('clearCoveredValues empties the covered cells and leaves the anchor alone', () => {
  const rows = grid({A1: 'anchor', B1: 'covered', A2: 'covered', C1: 'outside'});
  clearCoveredValues(rows, {top: 1, left: 1, bottom: 2, right: 2});
  assert.deepEqual(valuesOf(rows), ['A1="anchor"', 'A2=null', 'B1=null', 'C1="outside"']);
});

test('clearCoveredValues clears the value only: a covered cell keeps its style', () => {
  // A border spanning a merged region rides the covered cells, and is legal. Only the value is the
  // geometry Excel opens with a repair prompt.
  const rows = grid({A1: 'anchor', B1: 'covered'});
  const covered = rows.get(1)!.get(2)!;
  covered.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  clearCoveredValues(rows, {top: 1, left: 1, bottom: 1, right: 2});
  assert.equal(covered.value, null);
  assert.deepEqual(covered.fill, {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}});
});

test('clearCoveredValues walks only the rows a region names', () => {
  const rows = grid({A1: 'anchor', A5: 'far below'});
  clearCoveredValues(rows, {top: 1, left: 1, bottom: 2, right: 2});
  assert.equal(rows.get(5)!.get(1)!.value, 'far below');
});

// A `<mergeCell>` costs whatever it declares, and what it declares is free to be the whole grid:
// 4,000 non-overlapping full-column merges are a few hundred KB of XML that compresses to nothing.
// The guarantee is that the collapse walks the cells that exist, not the area claimed, so the counter
// aborts instead of the test hanging when a rewrite reintroduces a walk over the rectangle.
test('clearCoveredValues costs the populated cells, not the declared area', () => {
  const budget = {steps: 0};
  const rows = new CountedMap<Map<number, Cell>>(budget);
  for (let row = 1; row <= 4; row++) {
    const cols = new CountedMap<Cell>(budget);
    for (let col = 1; col <= 4; col++) cols.set(col, new Cell(row, col));
    rows.set(row, cols);
  }
  for (let col = 1; col <= 4000; col++) {
    clearCoveredValues(rows, {top: 1, left: col, bottom: 1_048_576, right: col});
  }
  assert.ok(budget.steps < 4000 * 40, `walked ${budget.steps} entries`);
});
