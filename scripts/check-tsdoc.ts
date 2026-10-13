#!/usr/bin/env node
// Every doc comment is TSDoc, as the TSDoc parser reads it.
//
// The comments were already TSDoc in spirit: no `{type}` annotations on `@param`, `{@link}` for
// references. What made them JSDoc was the handful of places a JSDoc habit leaked through, and
// nothing caught them: `@throws {RangeError}`, a `@param` with no hyphen, a `@param options.key`
// for one field of an object, a `<`, `{` or `}` in prose that a TSDoc reader takes for markup. A
// convention nobody checks drifts back to whatever the last author was used to, so the parser
// Microsoft maintains as the reference implementation is the check.
//
// "Doc comment" means what TypeScript attaches to a declaration, read from the compiler's own AST,
// rather than every `/**` a regex can find: a `/**` inside a string or a regex is not a comment,
// and a block attached to nothing documents nothing. The tags this repo adds to the standard set are
// declared in `tsdoc.json`, the file other TSDoc tools read for the same purpose, so that an editor
// plugin and this gate cannot disagree about what is defined.
//
// See docs/decisions/0044-doc-comments-are-tsdoc.md.
//
//   node scripts/check-tsdoc.ts

import {readFileSync} from 'node:fs';
import {join, relative} from 'node:path';

import {
  TSDocConfiguration,
  TSDocParser,
  TSDocTagDefinition,
  TSDocTagSyntaxKind,
  TextRange,
} from '@microsoft/tsdoc';
import type * as ast from 'typescript/unstable/ast';
import {API} from 'typescript/unstable/sync';

import {toPosix} from './module-graph.ts';
import {ROOT} from './repo.ts';
import {reportCrash, verdict} from './verdict.ts';

const GATE = 'tsdoc:check';

// Every project the typecheck gates open, so every authored `.ts` file is read by exactly the
// compiler options that check it.
const CONFIGS = ['tsconfig.json', 'tsconfig.test.json', 'www/tsconfig.json'].map((c) =>
  join(ROOT, c),
);

/** The part of `tsdoc.json` this reads: the tags this repo defines beyond the standard set. */
interface TsdocJson {
  readonly tagDefinitions?: readonly {
    readonly tagName: string;
    readonly syntaxKind: 'block' | 'inline' | 'modifier';
  }[];
}

const SYNTAX_KIND = {
  block: TSDocTagSyntaxKind.BlockTag,
  inline: TSDocTagSyntaxKind.InlineTag,
  modifier: TSDocTagSyntaxKind.ModifierTag,
} as const;

function configuration(): TSDocConfiguration {
  const config = new TSDocConfiguration();
  const json = JSON.parse(readFileSync(join(ROOT, 'tsdoc.json'), 'utf8')) as TsdocJson;
  for (const {tagName, syntaxKind} of json.tagDefinitions ?? []) {
    config.addTagDefinition(new TSDocTagDefinition({tagName, syntaxKind: SYNTAX_KIND[syntaxKind]}));
  }
  // An unknown tag is a typo or a JSDoc-ism until `tsdoc.json` says otherwise.
  config.validation.ignoreUndefinedTags = false;
  return config;
}

/** Every doc comment the compiler attaches to a node in this file, as `[pos, end)` ranges. */
function docComments(sourceFile: ast.SourceFile): {pos: number; end: number}[] {
  const ranges = new Map<number, number>();
  const visit = (node: ast.Node): void => {
    // A JSDoc node's `pos` is its full start, leading trivia included, so the comment itself opens
    // at the first `/**` after it.
    for (const doc of node.jsDoc ?? [])
      ranges.set(sourceFile.text.indexOf('/**', doc.pos), doc.end);
    node.forEachChild(visit);
  };
  visit(sourceFile);
  return [...ranges].map(([pos, end]) => ({pos, end}));
}

function lineOf(text: string, offset: number): number {
  let line = 1;
  for (let i = 0; i < offset; i++) if (text.charCodeAt(i) === 0x0a) line++;
  return line;
}

function main(): void {
  const parser = new TSDocParser(configuration());
  const problems: string[] = [];
  const seen = new Set<string>();
  let comments = 0;

  const api = new API({cwd: ROOT});
  try {
    const snapshot = api.updateSnapshot({openProjects: CONFIGS});
    try {
      for (const config of CONFIGS) {
        const project = snapshot.getProject(config);
        if (!project) throw new Error(`cannot open project ${config}`);
        for (const fileName of project.program.getSourceFileNames()) {
          const file = toPosix(relative(ROOT, fileName));
          if (file.startsWith('..') || file.includes('node_modules/') || seen.has(file)) continue;
          seen.add(file);
          const sourceFile = project.program.getSourceFile(fileName);
          if (!sourceFile) throw new Error(`cannot load ${file}`);
          const {text} = sourceFile;
          for (const {pos, end} of docComments(sourceFile)) {
            comments++;
            // A range over the whole file rather than a slice of it, so each message's offset is
            // already a file offset.
            const {log} = parser.parseRange(TextRange.fromStringRange(text, pos, end));
            for (const message of log.messages) {
              const line = lineOf(text, message.textRange.pos);
              problems.push(`${file}:${line}: ${message.unformattedText} [${message.messageId}]`);
            }
          }
        }
      }
    } finally {
      snapshot.dispose();
    }
  } finally {
    api.close();
  }

  verdict({
    gate: GATE,
    problems,
    ok: `${comments} doc comment(s) across ${seen.size} file(s) parse as TSDoc.`,
    failure: 'TSDoc problem(s)',
  });
}

try {
  main();
} catch (err) {
  reportCrash(GATE, err);
}
