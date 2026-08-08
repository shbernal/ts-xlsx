// Bundle-size budgets for the publishable build.
//
// Two numbers, because they answer two different questions.
//
// The **total** is every emitted `dist/**/*.js` — what ships in the tarball, and the tripwire
// against accidental bloat (an errant dependency inlined, dead code shipped).
//
// The **per-entry** numbers are what a consumer actually loads. Each public subpath in
// `package.json`'s `exports` is walked transitively through its static imports; the closure is the
// set of modules that must be present for that entry to evaluate. It is a lower bound on any
// bundler's answer — `sideEffects: false` lets a bundler prune *within* these modules, never add
// to them — and it is the only number that notices the failures that matter here: a codec
// acquiring a value-import of something it previously needed only as a type, or the model reaching
// into a parser. The total cannot see either; both leave it unchanged.
//
// Budgets are tripwires, not targets. Raise one deliberately, with the same eyes a dependency
// addition would get — and when you do, say in the commit *what* the entry gained.
//
//   node scripts/size-budget.ts

import {readdirSync, readFileSync, statSync} from 'node:fs';
import {dirname, join, normalize, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DIST = join(ROOT, 'dist');
const TOTAL_BUDGET_BYTES = 530 * 1024;

// Roughly a tenth of headroom over the measured closure, per entry: enough that ordinary growth is
// not a chore, tight enough that a whole codec crossing a boundary cannot hide inside it.
//
// Every number below was halved when `build` split into two tsc passes and the JS pass started
// stripping comments. That is not a budget cut — nothing left the closure, and no consumer loads a
// byte less than they did before the prose was measured as part of it. It is the measurement
// finally being of code. The old figures were ~47% comment, which is what had made this tripwire
// soft: a codec crossing a boundary is the failure these numbers exist to catch, and at the old
// scale one could have arrived inside a release's ordinary comment churn without moving them.
const ENTRY_BUDGETS_KB: Readonly<Record<string, number>> = {
  '.': 530,
  './core': 185,
  './xlsx': 520,
  './xlsb': 262,
  './csv': 190,
  './vba': 50,
  './customui': 14,
  // The taxonomy reaches nothing but itself, and that is the point: classifying a failure must
  // not cost a parser. A jump here means an error class started importing the layer it describes.
  './errors': 4,
};

interface PackageJson {
  readonly exports: Readonly<Record<string, string | {readonly default?: string}>>;
}

function jsFiles(dir: string): string[] {
  return readdirSync(dir, {withFileTypes: true}).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) return jsFiles(path);
    return entry.name.endsWith('.js') ? [path] : [];
  });
}

// Emitted JS, not source: the emitter picks its own quoting — double under TypeScript 6, single
// under 7 — so both forms are matched rather than any one being assumed.
const RELATIVE_SPECIFIER = /\bfrom\s+["'](\.[^"']*)["']/g;

/** Every relative specifier the emitted module imports or re-exports from, resolved to a path. */
function staticImports(file: string): string[] {
  const source = readFileSync(file, 'utf8');
  return [...source.matchAll(RELATIVE_SPECIFIER)].map((match) =>
    normalize(join(dirname(file), match[1] as string)),
  );
}

/** The modules that must be present for `entry` to evaluate, itself included. */
function closure(entry: string): Set<string> {
  const reached = new Set<string>();
  const pending = [entry];
  while (pending.length > 0) {
    const file = pending.pop() as string;
    if (reached.has(file)) continue;
    reached.add(file);
    pending.push(...staticImports(file));
  }
  return reached;
}

function bytes(files: Iterable<string>): number {
  let total = 0;
  for (const file of files) total += statSync(file).size;
  return total;
}

const kb = (n: number) => `${(n / 1024).toFixed(1)} KB`;

const pkg = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')) as PackageJson;
const over: string[] = [];

const all = jsFiles(DIST);
const total = bytes(all);
console.log(
  `total runtime JS: ${kb(total)} across ${all.length} file(s) — budget ${kb(TOTAL_BUDGET_BYTES)}`,
);
if (total > TOTAL_BUDGET_BYTES) {
  over.push(`total is over by ${kb(total - TOTAL_BUDGET_BYTES)}`);
}

console.log('\nper entry — the module closure a consumer of that subpath loads:\n');
for (const [subpath, target] of Object.entries(pkg.exports)) {
  const emitted = typeof target === 'string' ? undefined : target.default;
  if (emitted === undefined || !emitted.endsWith('.js')) continue;

  const budgetKb = ENTRY_BUDGETS_KB[subpath];
  if (budgetKb === undefined) {
    over.push(`"${subpath}" is published with no budget in ENTRY_BUDGETS_KB`);
    continue;
  }
  const reached = closure(resolve(ROOT, emitted));
  const size = bytes(reached);
  const budget = budgetKb * 1024;
  const verdict = size > budget ? `OVER by ${kb(size - budget)}` : 'ok';
  console.log(
    `  ${subpath.padEnd(12)} ${kb(size).padStart(9)}  ${String(reached.size).padStart(3)} modules   budget ${kb(budget).padStart(9)}   ${verdict}`,
  );
  if (size > budget) over.push(`"${subpath}" is over by ${kb(size - budget)}`);
}

if (over.length > 0) {
  console.error(`\nOVER BUDGET:\n${over.map((line) => `  ${line}`).join('\n')}`);
  console.error('\nInvestigate the growth or raise the budget deliberately.');
  process.exit(1);
}
