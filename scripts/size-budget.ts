// Bundle-size budgets for the publishable build.
//
// Two numbers, because they answer two different questions.
//
// The **total** is every emitted `dist/**/*.js`: what ships in the tarball, and the tripwire
// against accidental bloat (an errant dependency inlined, dead code shipped).
//
// The **per-entry** numbers are what a consumer actually loads. Each public subpath in
// `package.json`'s `exports` is walked transitively through its static imports; the closure is the
// set of modules that must be present for that entry to evaluate. It is a lower bound on any
// bundler's answer, since `sideEffects: false` lets a bundler prune *within* these modules and
// never add to them. It is also the only number that notices the failures that matter here:
// a codec acquiring a value-import of something it previously needed only as a type, or the model
// reaching into a parser. The total cannot see either; both leave it unchanged.
//
// Budgets are tripwires, not targets. Raise one deliberately, with the same eyes a dependency
// addition would get. When you do, say in the commit *what* the entry gained.
//
// Run by the `size` gate of `verify --full`, which builds first.
// It used to be the publish step's alone, and `/customui` spent a release 3 KB over its budget on a
// tree that was green everywhere anyone looked: a budget only the publish step checks is not a
// tripwire. `docs/architecture.md` ("The size budgets are per entry") carries the reasoning.
//
//   node scripts/size-budget.ts

import {statSync} from 'node:fs';
import {join, resolve} from 'node:path';

import {closure, importedPaths, sourceFiles} from './module-graph.ts';
import {readPackageJson, ROOT} from './repo.ts';
import {verdict} from './verdict.ts';

const DIST = join(ROOT, 'dist');
const TOTAL_BUDGET_BYTES = 685 * 1024;

// Roughly a tenth of headroom over the measured closure, per entry: enough that ordinary growth is
// not a chore, tight enough that a whole codec crossing a boundary cannot hide inside it.
const ENTRY_BUDGETS_KB: Readonly<Record<string, number>> = {
  '.': 673,
  './core': 248,
  './xlsx': 655,
  './xlsb': 360,
  './csv': 259,
  // The streaming writer and the write half it rides on, and nothing of the reader: a jump here is
  // the read path arriving, which would mean the entry had stopped being about one thing.
  './node': 494,
  './vba': 57,
  './customui': 16,
  // The taxonomy reaches nothing but itself, and that is the point: classifying a failure must
  // not cost a parser. A jump here means an error class started importing the layer it describes.
  './errors': 4,
};

function bytes(files: Iterable<string>): number {
  let total = 0;
  for (const file of files) total += statSync(file).size;
  return total;
}

const kb = (n: number) => `${(n / 1024).toFixed(1)} KB`;

const pkg = readPackageJson();
const over: string[] = [];

const all = sourceFiles(DIST, '.js');
const total = bytes(all);
console.log(
  `total runtime JS: ${kb(total)} across ${all.length} file(s); budget ${kb(TOTAL_BUDGET_BYTES)}`,
);
if (total > TOTAL_BUDGET_BYTES) {
  over.push(`total is over by ${kb(total - TOTAL_BUDGET_BYTES)}`);
}

console.log('\nper entry, the module closure a consumer of that subpath loads:\n');
for (const [subpath, target] of Object.entries(pkg.exports)) {
  const emitted = typeof target === 'string' ? undefined : target.default;
  if (emitted === undefined || !emitted.endsWith('.js')) continue;

  const budgetKb = ENTRY_BUDGETS_KB[subpath];
  if (budgetKb === undefined) {
    over.push(`"${subpath}" is published with no budget in ENTRY_BUDGETS_KB`);
    continue;
  }
  const reached = closure(resolve(ROOT, emitted), importedPaths);
  const size = bytes(reached);
  const budget = budgetKb * 1024;
  const state = size > budget ? `OVER by ${kb(size - budget)}` : 'ok';
  console.log(
    `  ${subpath.padEnd(12)} ${kb(size).padStart(9)}  ${String(reached.size).padStart(3)} modules   budget ${kb(budget).padStart(9)}   ${state}`,
  );
  if (size > budget) over.push(`"${subpath}" is over by ${kb(size - budget)}`);
}

// Through `verdict` like every sibling, and `process.exitCode` rather than `process.exit`. This is
// a publish-blocking gate whose output `verify.ts` captures through a pipe, which is the exact
// configuration `verdict.ts`'s header describes: `process.exit` truncates an unflushed pipe, so the
// gate that most needs its diagnostic read was the one that could lose it.
verdict({
  gate: 'size',
  problems: over.map((line) => `  ${line}`),
  ok: 'every entry within its budget',
  failure: 'entry closure(s) over budget; investigate the growth or raise the budget deliberately',
});
