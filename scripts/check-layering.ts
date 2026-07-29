#!/usr/bin/env node
// The module graph's direction, enforced.
//
// The shared container layer (OPC/ZIP, relationship resolution, format sniffing) and the XML
// helpers used to live inside `src/io/xlsx/`, so the BIFF12 codec reached sideways into the XML
// codec to get at them. That is a direction no codec should have over another, and it regresses in
// one careless import — an import that typechecks, passes every test, and looks locally reasonable.
// A comment asking the next author to remember is not a mechanism; this is.
//
// Each rule names a directory and the directories its modules may not import. Test files are
// exempt: a co-located white-box test may reach anywhere in src (see the test-topology rule in
// docs/architecture.md), and a test import creates no dependency in the graph we ship.
//
//   node scripts/check-layering.ts

import {readdirSync, readFileSync, statSync} from 'node:fs';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

interface Rule {
  /** Modules under this directory — or the single module at this exact path… */
  readonly layer: string;
  /** …may not import modules under any of these. */
  readonly forbidden: readonly string[];
  /** Why, in one line — printed with the violation so the fix is obvious from the failure alone. */
  readonly because: string;
}

const RULES: readonly Rule[] = [
  {
    layer: 'src/errors.ts',
    forbidden: ['src/core', 'src/io', 'src/xml', 'src/vba', 'src/customui'],
    because: 'the failure taxonomy is below every layer that throws through it',
  },
  {
    layer: 'src/xml',
    forbidden: ['src/core', 'src/io', 'src/vba', 'src/customui'],
    because: 'XML escaping and parsing know nothing about spreadsheets',
  },
  {
    layer: 'src/core',
    // Not `src/vba`/`src/customui`: a workbook models a VBA project and its ribbon parts, so those
    // types are part of the document, not of a way to write it down.
    forbidden: ['src/io', 'src/xml'],
    because: 'the model does not know which serialisations exist',
  },
  {
    layer: 'src/io/opc',
    forbidden: ['src/io/xlsx', 'src/io/xlsb', 'src/io/csv', 'src/io/style'],
    because: 'the container layer is below every codec that rides in it',
  },
  {
    layer: 'src/io/style',
    forbidden: ['src/io/xlsx', 'src/io/xlsb', 'src/io/csv'],
    because: 'the resolved-format model is shared by the codecs, not owned by one',
  },
  {
    layer: 'src/io/xlsb',
    forbidden: ['src/io/xlsx'],
    because:
      'the BIFF12 and XML codecs are peers — shared code belongs in src/io/opc or src/io/style',
  },
];

/**
 * The entry barrels are the package's public faces, not modules to build on. Only the root barrel
 * composes them; an internal import of one would route a dependency through the public surface —
 * the module graph's shape would then follow what we chose to export, and a cycle would be one
 * re-export away. Their own contents are checked by `scripts/check-entries.ts`.
 */
const ENTRIES = 'src/entries';
const ENTRY_COMPOSER = 'src/index.ts';

function sourceFiles(dir: string): string[] {
  return readdirSync(dir).flatMap((name) => {
    const path = `${dir}/${name}`;
    if (statSync(join(ROOT, path)).isDirectory()) return sourceFiles(path);
    return path.endsWith('.ts') && !path.endsWith('.test.ts') ? [path] : [];
  });
}

// Every relative specifier the module imports or re-exports from, resolved to a repo-relative path.
// Only relative specifiers can cross a layer — a bare specifier is a dependency, not a layer.
function importedPaths(file: string): string[] {
  const source = readFileSync(join(ROOT, file), 'utf8');
  // Both spellings: `… from '…'`, and the bare `import '…'` that has no `from` to anchor on. The
  // bare form should never appear — the package declares `"sideEffects": false`, so an import kept
  // only for its effect is a lie to every bundler — but a rule that cannot see it would say the
  // graph is clean while a layer was being crossed by the one import form it was blind to.
  const specifiers = [...source.matchAll(/\b(?:from|import)\s+'(\.[^']*)'/g)].map(
    (match) => match[1] as string,
  );
  const dir = file.slice(0, file.lastIndexOf('/'));
  return specifiers.map((specifier) => {
    const segments = `${dir}/${specifier}`.split('/');
    const out: string[] = [];
    for (const segment of segments) {
      if (segment === '' || segment === '.') continue;
      if (segment === '..') out.pop();
      else out.push(segment);
    }
    return out.join('/');
  });
}

const violations: string[] = [];
for (const file of sourceFiles('src')) {
  if (file !== ENTRY_COMPOSER && !file.startsWith(`${ENTRIES}/`)) {
    for (const target of importedPaths(file)) {
      if (target.startsWith(`${ENTRIES}/`)) {
        violations.push(
          `  ${file}\n    imports ${target}\n    only ${ENTRY_COMPOSER} may compose the entry barrels — import the module that declares the symbol`,
        );
      }
    }
  }
  const rule = RULES.find(
    (candidate) => file === candidate.layer || file.startsWith(`${candidate.layer}/`),
  );
  if (rule === undefined) continue;
  for (const target of importedPaths(file)) {
    const crossed = rule.forbidden.find((layer) => target.startsWith(`${layer}/`));
    if (crossed !== undefined) {
      violations.push(
        `  ${file}\n    imports ${target}\n    ${rule.layer} may not reach into ${crossed} — ${rule.because}`,
      );
    }
  }
}

if (violations.length === 0) {
  console.log(`layering: ${RULES.length} rules + the entry-barrel rule hold across src/`);
} else {
  console.error(`\nlayering: ${violations.length} import(s) cross a layer boundary.\n`);
  console.error(`${violations.join('\n\n')}\n`);
  process.exit(1);
}
