#!/usr/bin/env node
// The exhaustiveness register lists every exhaustiveness proof.
//
// `src/type-tests/facet-tables.type-test.ts` exists to state, in one place, which facet tables carry
// an `AssertNever` guarantee, so a table added without one reads as an absence rather than being
// assumed. That is worth exactly as much as the list is complete, and it had drifted to six of
// eleven: `EveryColumnPropertyIsMirrored` was listed and its literal twin `EveryRowPropertyIsMirrored`
// was not, and all four of `page-setup.ts`'s proofs were missing.
//
// Nothing was unproven while it drifted, because `AssertNever<T extends never>` fails at the
// declaration. What was wrong was the register's own claim: read it and you concluded `RowProperties`
// and `PageSetup` carried no proof at all. This holds the two in step, so the next one added is
// either listed or a failed gate.
//
//   node scripts/check-facet-register.ts

import {readFileSync} from 'node:fs';
import {join} from 'node:path';

import {sourceFiles, toPosix} from './module-graph.ts';
import {ROOT} from './repo.ts';
import {verdict} from './verdict.ts';

const REGISTER = 'src/type-tests/facet-tables.type-test.ts';

// The shape every proof takes: an exported alias resolving to `never` when the table is complete.
const PROOF = /^export type (\w+) = AssertNever</gm;

function declaredProofs(): Map<string, string> {
  const proofs = new Map<string, string>();
  for (const absolute of sourceFiles(join(ROOT, 'src'), '.ts')) {
    const file = toPosix(absolute).slice(toPosix(ROOT).length + 1);
    if (file === REGISTER) continue;
    const text = readFileSync(absolute, 'utf8');
    PROOF.lastIndex = 0;
    for (let match = PROOF.exec(text); match !== null; match = PROOF.exec(text)) {
      const name = match[1];
      if (name !== undefined) proofs.set(name, file);
    }
  }
  return proofs;
}

const register = readFileSync(join(ROOT, REGISTER), 'utf8');
const proofs = declaredProofs();
const problems: string[] = [];
for (const [name, file] of proofs) {
  // Asserted, not merely imported: an import the list does not spend is a name in scope and no claim.
  if (!register.includes(`Expect<Equal<${name}, never>>`)) {
    problems.push(
      `  ${name} (${file})\n    is an exhaustiveness proof that ${REGISTER} does not assert:\n` +
        `    add Expect<Equal<${name}, never>> so the register still names every table that carries one`,
    );
  }
}

verdict({
  gate: 'facet register',
  problems,
  ok: `${proofs.size} exhaustiveness proof(s), every one asserted in the register`,
  failure: 'exhaustiveness proof(s) missing from the register',
});
