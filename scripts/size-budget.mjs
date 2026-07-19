// Bundle-size budget for the publishable build.
//
// The meaningful number is the total emitted runtime JavaScript — what a consumer
// actually loads — not the tarball (which varies with maps/metadata). We sum every
// dist/**/*.js and fail if it crosses the budget. The budget is a tripwire against
// accidental bloat (an errant dependency inlined, dead code shipped), not a target;
// raise it deliberately, with the same eyes a dependency addition would get.

import {readdirSync, statSync} from 'node:fs';
import {join} from 'node:path';

const DIST = 'dist';
const BUDGET_BYTES = 600 * 1024;

function jsFiles(dir) {
  const out = [];
  for (const entry of readdirSync(dir, {withFileTypes: true})) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...jsFiles(path));
    else if (entry.name.endsWith('.js')) out.push(path);
  }
  return out;
}

let total = 0;
let count = 0;
for (const file of jsFiles(DIST)) {
  total += statSync(file).size;
  count += 1;
}

const kb = (n) => `${(n / 1024).toFixed(1)} KB`;
console.log(`runtime JS: ${kb(total)} across ${count} file(s) — budget ${kb(BUDGET_BYTES)}`);

if (total > BUDGET_BYTES) {
  console.error(`\nOVER BUDGET by ${kb(total - BUDGET_BYTES)}. Investigate the growth or raise the budget deliberately.`);
  process.exit(1);
}
