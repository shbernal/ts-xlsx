#!/usr/bin/env node

/**
 * The `ooxml-lookup` CLI.
 *
 * Runs from a bare checkout: no package.json, no install, no node_modules.
 * Node 24+ and nothing else. Every module it imports is either a `node:`
 * builtin or a relative path, and that is a property `core/skill-cli.test.mjs`
 * asserts rather than trusts.
 *
 * **Everything it prints on stdout is JSON.** That is the contract, and it is
 * why the experimental-warning filter is load-bearing here rather than
 * cosmetic: without it `node:sqlite` writes a paragraph of English to stderr on
 * the first call, and an agent piping this into a parser sees noise.
 */
import {explainDiagnostic} from './explain.mjs';
import {createGraph} from './graph.mjs';

const USAGE = `ooxml — query the ECMA-376 schema graph, offline.

  ooxml element    <qname>      what a name is: kind, type, namespace, profiles
  ooxml type       <qname>      a complexType/simpleType: derivation and shape
  ooxml children   <qname>      legal children, in order, with cardinality
  ooxml attributes <qname>      attributes, inherited and attributeGroups expanded
  ooxml enum       <qname>      enumeration values of a simple type
  ooxml values     <qname>      the whole value space: facets, patterns, unions
  ooxml namespace  <uri|prefix> namespace <-> vocabulary, both directions
  ooxml search     <substring>  find symbols by name (substring, not semantic)
  ooxml diff       <qname>      what Transitional adds to Strict for this symbol
  ooxml explain    <json>       resolve a validator diagnostic against the schema
  ooxml sql        <select>     read-only SQL against the graph (advanced)

Options:
  --profile <transitional|strict>   default: transitional
  --limit <n>                       search/sql only
  --compact                         force compact JSON (default when not a TTY)

Names may be written w:tblPr, wml:tblPr, {namespace-uri}tblPr, or bare.
A bare name that means several things returns all of them.`;

/** `--flag value` and `--flag=value`, with everything else positional. */
function parseArgv(argv) {
  const positional = [];
  const options = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith('--')) {
      positional.push(arg);
      continue;
    }
    const eq = arg.indexOf('=');
    if (eq !== -1) {
      options[arg.slice(2, eq)] = arg.slice(eq + 1);
    } else if (arg === '--compact' || arg === '--help') {
      options[arg.slice(2)] = true;
    } else {
      options[arg.slice(2)] = argv[++i];
    }
  }
  return {positional, options};
}

function main(argv) {
  const {positional, options} = parseArgv(argv);
  const [command, ...rest] = positional;

  if (command === undefined || options.help === true || command === 'help') {
    process.stdout.write(`${USAGE}\n`);
    return 0;
  }

  const profile = options.profile ?? 'transitional';
  // Indentation is for a human at a terminal. Piped into an agent's context it
  // is pure cost, and this tool is used in a loop.
  const indent = options.compact === true || !process.stdout.isTTY ? 0 : 2;
  const emit = (value) => process.stdout.write(`${JSON.stringify(value, null, indent)}\n`);

  const graph = createGraph();
  try {
    const argument = rest.join(' ');
    const need = (what) => {
      if (argument === '') throw new UsageError(`${command} needs ${what}`);
      return argument;
    };

    switch (command) {
      case 'element':
        emit(graph.element(need('a name'), {profile}));
        return 0;
      case 'type':
        emit(graph.type(need('a name'), {profile}));
        return 0;
      case 'children':
        emit(graph.children(need('a name'), {profile}));
        return 0;
      case 'attributes':
        emit(graph.attributes(need('a name'), {profile}));
        return 0;
      case 'enum':
        emit(graph.enum(need('a name'), {profile}));
        return 0;
      case 'values':
        emit(graph.values(need('a name'), {profile}));
        return 0;
      case 'namespace':
        emit(graph.namespace(need('a URI, prefix or vocabulary')));
        return 0;
      case 'search':
        emit(
          graph.search(need('a substring'), {
            profile,
            limit: options.limit === undefined ? 40 : Number(options.limit),
          }),
        );
        return 0;
      case 'diff':
      case 'diff-profiles':
        emit(graph.diff_profiles(need('a name')));
        return 0;
      case 'explain':
        emit(explainDiagnostic(graph, parseDiagnostic(need('a diagnostic as JSON')), {profile}));
        return 0;
      case 'sql':
        emit(runSql(graph, need('a SELECT statement'), options.limit));
        return 0;
      default:
        throw new UsageError(`unknown command "${command}"`);
    }
  } finally {
    graph.close();
  }
}

class UsageError extends Error {}

/**
 * The diagnostic, as JSON. Accepts a whole `ooxml-validate` report and takes its
 * first diagnostic, because that is what people actually have in hand — but the
 * four fields are still all that is read (see core/explain.mjs).
 */
function parseDiagnostic(text) {
  let value;
  try {
    value = JSON.parse(text);
  } catch {
    throw new UsageError(
      'explain expects JSON. Pass the diagnostic object from an ooxml-validate ' +
        'report, e.g. \'{"id":"Sch_UndeclaredAttribute","description":"...","xpath":"/w:document[1]"}\'',
    );
  }
  if (Array.isArray(value)) return value[0] ?? {};
  if (Array.isArray(value?.diagnostics)) return value.diagnostics[0] ?? {};
  return value;
}

/**
 * Read-only SQL, the escape hatch the subcommands do not cover.
 *
 * This is the skill's real advantage over the MCP surface: the agent has a
 * shell and the database is right there, so a documented schema plus a query
 * runner beats any fixed set of subcommands for the question nobody anticipated.
 *
 * The connection is opened read-only, so this cannot write whatever it is
 * asked; the statement check below is about giving a clear error rather than a
 * SQLite one, and about not pretending a multi-statement script will work.
 */
function runSql(graph, statement, limit) {
  const trimmed = statement.trim().replace(/;\s*$/, '');
  if (!/^(select|with)\b/i.test(trimmed)) {
    throw new UsageError(
      'sql accepts a single SELECT (or WITH ... SELECT); the database is read-only',
    );
  }
  if (trimmed.includes(';')) {
    throw new UsageError('sql runs one statement at a time');
  }

  const max = limit === undefined ? 200 : Number(limit);
  const rows = graph._internal.handle.all(trimmed);
  return {
    sql: trimmed,
    count: rows.length,
    // Truncation is reported rather than silent — a quietly cut result set is
    // indistinguishable from a complete one.
    truncated: rows.length > max,
    rows: rows.slice(0, max),
  };
}

try {
  process.exitCode = main(process.argv.slice(2));
} catch (error) {
  const usage = error instanceof UsageError;
  process.stderr.write(`${error.message}\n${usage ? `\n${USAGE}\n` : ''}`);
  process.exitCode = usage ? 2 : 1;
}
