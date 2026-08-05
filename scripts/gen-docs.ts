// Generates the public API reference under `docs/api/` straight from the types.
//
// The public barrel (`src/index.ts`) is the single source of truth: this walks the
// symbols it re-exports via the TypeScript compiler API, renders each one's JSDoc
// summary + tags + a body-stripped TypeScript signature, and writes one Markdown
// page per originating module plus an index. No new dependency — `typescript` is
// already the toolchain — so the docs cannot describe a shape the compiler wouldn't
// accept. Run `pnpm run docs`; `pnpm run docs:check` fails if the committed pages have
// drifted from a fresh generation (the docs are gated like any other artifact).
//
// See docs/decisions/0006-docs-from-types.md.

import {mkdirSync, rmSync, writeFileSync} from 'node:fs';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import ts from 'typescript';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const ENTRY = join(ROOT, 'src/index.ts');
const OUT_DIR = join(ROOT, 'docs/api');

// Human-facing page titles + ordering, keyed by the source module basename an export
// resolves to. A module not listed here still renders (alphabetically, titled from its
// filename) so a new public module can never silently vanish from the docs.
const GROUPS: ReadonlyArray<readonly [string, string]> = [
  ['address', 'Addresses & ranges'],
  ['value', 'Cell values'],
  ['cell', 'Cell'],
  ['workbook', 'Workbook'],
  ['worksheet', 'Worksheet'],
  ['style', 'Styles'],
  ['table', 'Tables'],
  ['pivot-table', 'Pivot tables'],
  ['image', 'Images'],
  ['protection', 'Protection'],
  ['workbook-protection', 'Protection'],
  ['xlsx/read', 'Reading .xlsx'],
  ['xlsx/read-rows', 'Streaming reads'],
  ['xlsx/write', 'Writing .xlsx'],
  ['xlsx/write-stream', 'Streaming writes'],
  ['csv/read', 'CSV'],
  ['csv/write', 'CSV'],
];

/** The body-bearing declarations — the ones whose signature ends where their body begins. */
function bodyOf(node: ts.Node): ts.Node | undefined {
  if (
    ts.isFunctionDeclaration(node) ||
    ts.isMethodDeclaration(node) ||
    ts.isConstructorDeclaration(node) ||
    ts.isGetAccessorDeclaration(node) ||
    ts.isSetAccessorDeclaration(node)
  ) {
    return node.body;
  }
  return undefined;
}

/**
 * Render a declaration as a signature by slicing its own source text, cut at the body when it
 * has one.
 *
 * Source text rather than a compiler re-print: the reference then shows the shape the author
 * wrote — parameter line breaks and all — instead of the printer's normalization, and the
 * generator needs no emit machinery. Starting at `getStart` drops leading trivia, so the JSDoc
 * above a declaration stays out of the code block that renders it.
 */
function printSignature(node: ts.Node, sourceFile: ts.SourceFile): string {
  const body = bodyOf(node);
  const text = sourceFile.text
    .slice(node.getStart(sourceFile), body ? body.getStart(sourceFile) : node.getEnd())
    .trimEnd()
    .replace(/^export (?:default )?/, '');
  // Cutting the body leaves the signature unterminated; `;` restores the `.d.ts` shape a reader
  // expects. Declarations that keep their text (interfaces, type aliases) already carry their own.
  return body ? `${text};` : text;
}

/**
 * Turn TSDoc `{@link Target}` / `{@link Target | label}` into a plain code span.
 *
 * Only the pipe form carries a label. TSDoc also permits `{@link Target label}` (space, no pipe), which
 * falls through here and renders as the whole string — so write the pipe, or write a bare target.
 */
function resolveLinks(text: string): string {
  return text.replace(
    /\{@link(?:code|plain)?\s+([^}|]+?)(?:\s*\|\s*([^}]+))?\}/g,
    (_m: string, target: string, label: string | undefined) => `\`${(label ?? target).trim()}\``,
  );
}

// A bare `@word` mid-sentence silently truncates the rendered summary: TypeScript's JSDoc parser reads
// it as the start of an unknown tag and drops everything after it. Prose about, say, an `@mention` has
// to wrap it in backticks — the summary comes from the compiler, so this cannot be fixed downstream.
function docText(symbol: ts.Symbol, checker: ts.TypeChecker): string {
  return resolveLinks(ts.displayPartsToString(symbol.getDocumentationComment(checker)).trim());
}

/** Render `@throws`, `@param`, `@returns`, `@example` tags into Markdown, in a stable order. */
function docTags(symbol: ts.Symbol, checker: ts.TypeChecker): string[] {
  const order: Record<string, number> = {param: 0, returns: 1, throws: 2, example: 3};
  const tags = symbol
    .getJsDocTags(checker)
    .filter((t) => t.name in order)
    .sort((a, b) => (order[a.name] ?? 0) - (order[b.name] ?? 0));
  const lines: string[] = [];
  for (const tag of tags) {
    const text = resolveLinks(ts.displayPartsToString(tag.text).trim());
    if (tag.name === 'example') {
      lines.push('', '```ts', text, '```');
    } else if (tag.name === 'param') {
      const [name, ...rest] = text.split(/\s+/);
      lines.push(`- \`${name}\` — ${rest.join(' ')}`);
    } else if (tag.name === 'returns') {
      lines.push(`**Returns** — ${text}`);
    } else if (tag.name === 'throws') {
      lines.push(`**Throws** — ${text.replace(/^\{[^}]*\}\s*/, '')}`);
    }
  }
  return lines;
}

function kindLabel(node: ts.Node): string {
  if (ts.isInterfaceDeclaration(node)) return 'interface';
  if (ts.isTypeAliasDeclaration(node)) return 'type';
  if (ts.isClassDeclaration(node)) return 'class';
  if (ts.isFunctionDeclaration(node)) return 'function';
  if (ts.isEnumDeclaration(node)) return 'enum';
  if (ts.isVariableDeclaration(node)) return 'const';
  return 'value';
}

function isPublicMember(member: ts.ClassElement): boolean {
  if (member.name && ts.isPrivateIdentifier(member.name)) return false;
  // A computed name is the codec's back channel (`src/core/internal.ts`), keyed by a symbol that
  // never leaves the package. Reachable only by a holder of that symbol, so it is no more public
  // than a `#private` field — and rendering it would dump the whole channel, initializer included,
  // into the reference a caller reads.
  if (member.name && ts.isComputedPropertyName(member.name)) return false;
  const mods = ts.canHaveModifiers(member) ? ts.getModifiers(member) : undefined;
  if (mods?.some((m) => m.kind === ts.SyntaxKind.PrivateKeyword)) return false;
  if (mods?.some((m) => m.kind === ts.SyntaxKind.ProtectedKeyword)) return false;
  return Boolean(member.name);
}

/**
 * A class renders as a signature overview (every public member) plus a prose list of
 * each documented member — overloads share one summary. Returns the code-block member
 * lines and the Markdown doc list separately so the reader gets a compact signature
 * block above readable per-member notes.
 */
function renderClassMembers(
  node: ts.ClassDeclaration,
  sourceFile: ts.SourceFile,
  checker: ts.TypeChecker,
): {sigs: string[]; docs: string[]} {
  const sigs: string[] = [];
  const docs: string[] = [];
  const documented = new Set<string>();
  for (const member of node.members) {
    if (!isPublicMember(member)) continue;
    if (!member.name) continue;
    const name = member.name.getText(sourceFile);
    const sig = printSignature(member, sourceFile).trim();
    sigs.push(`  ${sig}`);
    if (documented.has(name)) continue;
    const sym = checker.getSymbolAtLocation(member.name);
    const summary = sym ? docText(sym, checker) : '';
    if (summary) {
      documented.add(name);
      // The overview block above renders a signature as the author wrote it, line breaks and
      // all; this list puts one in an inline code span, which cannot survive them.
      docs.push(`- \`${sig.replace(/\s+/g, ' ')}\` — ${summary.replace(/\s+/g, ' ')}`);
    }
  }
  return {sigs, docs};
}

function main() {
  const program = ts.createProgram([ENTRY], {
    module: ts.ModuleKind.NodeNext,
    moduleResolution: ts.ModuleResolutionKind.NodeNext,
    target: ts.ScriptTarget.ES2022,
    allowImportingTsExtensions: true,
    noEmit: true,
    strict: true,
  });
  const checker = program.getTypeChecker();
  const entrySf = program.getSourceFile(ENTRY);
  if (!entrySf) throw new Error(`cannot load entry ${ENTRY}`);
  const moduleSymbol = checker.getSymbolAtLocation(entrySf);
  if (!moduleSymbol) throw new Error('entry has no module symbol — is src/index.ts a module?');

  const groupTitle = new Map(GROUPS);
  const groupOrder = new Map<string, number>(GROUPS.map(([key], i) => [key, i]));
  type Entry = {name: string; block: string};
  type Page = {title: string; key: string; entries: Entry[]};
  const pages = new Map<string, Page>();

  for (const exported of checker.getExportsOfModule(moduleSymbol)) {
    const symbol =
      exported.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(exported) : exported;
    const decl = symbol.getDeclarations()?.[0];
    if (!decl) continue;
    const sourceFile = decl.getSourceFile();
    const rel = sourceFile.fileName
      .replace(/^.*\/src\//, '')
      .replace(/\.ts$/, '')
      .replace(/\/index$/, '');
    const groupKey = rel.replace(/^core\//, '').replace(/^io\//, '');
    const title = groupTitle.get(groupKey) ?? titleize(groupKey);
    const pageId = title;

    let page = pages.get(pageId);
    if (!page) {
      page = {title, key: groupKey, entries: []};
      pages.set(pageId, page);
    }

    const name = exported.getName();
    const kind = kindLabel(decl);
    const summary = docText(symbol, checker);
    const tagLines = docTags(symbol, checker);

    const block = [`### \`${name}\``, '', `<sub>${kind}</sub>`, ''];
    if (summary) block.push(summary, '');

    let signature: string;
    let memberDocs: string[] = [];
    if (ts.isClassDeclaration(decl)) {
      const {sigs, docs} = renderClassMembers(decl, sourceFile, checker);
      const heritage = decl.heritageClauses?.map((h) => h.getText(sourceFile)).join(' ');
      signature = `class ${name}${heritage ? ` ${heritage}` : ''} {\n${sigs.join('\n')}\n}`;
      memberDocs = docs;
    } else if (ts.isVariableDeclaration(decl)) {
      const type = checker.typeToString(
        checker.getTypeOfSymbolAtLocation(symbol, decl),
        decl,
        ts.TypeFormatFlags.NoTruncation,
      );
      signature = `const ${name}: ${type}`;
    } else if (ts.isFunctionDeclaration(decl)) {
      // Print every overload declaration (those without a body); skip the impl signature.
      const declarations = symbol.getDeclarations() ?? [];
      const overloads = declarations.filter(
        (d): d is ts.FunctionDeclaration =>
          ts.isFunctionDeclaration(d) && (!d.body || declarations.length === 1),
      );
      signature = overloads.map((d) => printSignature(d, d.getSourceFile())).join('\n');
    } else {
      signature = printSignature(decl, sourceFile);
    }
    block.push('```ts', signature, '```');
    if (tagLines.length > 0) block.push('', ...tagLines);
    if (memberDocs.length > 0) block.push('', '**Members**', '', ...memberDocs);

    page.entries.push({name, block: block.join('\n')});
  }

  // Deterministic output regardless of `getExportsOfModule` iteration order — the CI
  // drift check compares committed pages against a fresh generation byte for byte.
  for (const page of pages.values()) {
    page.entries.sort((a: Entry, b: Entry) => a.name.localeCompare(b.name));
  }

  const ordered = [...pages.values()].sort((a, b) => {
    const oa = groupOrder.get(a.key) ?? Number.MAX_SAFE_INTEGER;
    const ob = groupOrder.get(b.key) ?? Number.MAX_SAFE_INTEGER;
    return oa - ob || a.title.localeCompare(b.title);
  });

  rmSync(OUT_DIR, {recursive: true, force: true});
  mkdirSync(OUT_DIR, {recursive: true});

  for (const page of ordered) {
    const slug = slugify(page.title);
    const body = [
      `# ${page.title}`,
      '',
      '<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->',
      '',
      page.entries.map((e) => e.block).join('\n\n---\n\n'),
      '',
    ].join('\n');
    writeFileSync(join(OUT_DIR, `${slug}.md`), body);
  }

  const index = [
    '# API reference',
    '',
    '<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->',
    '',
    'Every symbol below is re-exported from the package root and generated straight from',
    'its TypeScript declaration — the types are the contract.',
    '',
    ...ordered.map((page) => {
      const slug = slugify(page.title);
      const symbols = page.entries
        .map((e) => `[\`${e.name}\`](./${slug}.md#${anchor(e.name)})`)
        .join(', ');
      return `- **[${page.title}](./${slug}.md)** — ${symbols}`;
    }),
    '',
  ].join('\n');
  writeFileSync(join(OUT_DIR, 'README.md'), index);

  const count = ordered.reduce((n, p) => n + p.entries.length, 0);
  process.stdout.write(`docs: ${count} symbols across ${ordered.length} pages → docs/api/\n`);
}

function titleize(key: string): string {
  return key.replace(/[-/]/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase());
}
function slugify(title: string): string {
  return title
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
}
function anchor(name: string): string {
  return name.toLowerCase().replace(/[^a-z0-9]+/g, '');
}

main();
