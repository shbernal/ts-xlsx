#!/usr/bin/env node
// Harvest a single upstream issue or PR into the local backlog dataset.
//
// Usage:  node scripts/harvest/fetch-issue.mjs <number> [--repo owner/name]
//
// This is the *atom* of Phase 1 harvesting: it turns one GitHub thread into a
// durable, queryable JSON record under docs/knowledge/backlog/, plus any
// spreadsheet-shaped attachments it can find. Phase 1 fans this out across the
// whole backlog; here it exists so the dataset format is proven end-to-end
// before we scale. See STRATEGY.md Phase 0/1 and docs/knowledge/backlog/README.md.
//
// Auth comes from the `gh` CLI (already logged in); we shell out to `gh api`
// rather than embedding a token so credentials never touch this process.

import {execFile} from 'node:child_process';
import {mkdir, writeFile} from 'node:fs/promises';
import {dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {promisify} from 'node:util';

const execFileAsync = promisify(execFile);

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..', '..');
const BACKLOG_DIR = resolve(REPO_ROOT, 'docs', 'knowledge', 'backlog');
const DEFAULT_REPO = 'exceljs/exceljs';

// A spreadsheet library ingests hostile input; the harvester ingests hostile
// URLs. Cap attachment downloads so a malicious/huge asset can't exhaust disk.
const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
const SPREADSHEET_EXTENSIONS = new Set(['xlsx', 'xlsm', 'xlsb', 'xls', 'csv', 'zip', 'ods']);

function parseArgs(argv) {
  const args = {number: undefined, repo: DEFAULT_REPO};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--repo') {
      args.repo = argv[++i];
    } else if (/^\d+$/.test(arg)) {
      args.number = Number(arg);
    } else {
      throw new Error(`Unrecognized argument: ${arg}`);
    }
  }
  if (!Number.isInteger(args.number) || args.number <= 0) {
    throw new Error('Expected a positive issue/PR number, e.g. `fetch-issue.mjs 140`');
  }
  if (!/^[\w.-]+\/[\w.-]+$/.test(args.repo)) {
    throw new Error(`Invalid --repo value: ${args.repo}`);
  }
  return args;
}

async function gh(path, {paginate = false} = {}) {
  const argv = ['api', path];
  if (paginate) argv.push('--paginate');
  const {stdout} = await execFileAsync('gh', argv, {maxBuffer: 64 * 1024 * 1024});
  if (paginate) {
    // `gh --paginate` concatenates JSON arrays as `][` at page seams; splice them.
    return JSON.parse(stdout.replace(/\]\s*\[/g, ','));
  }
  return JSON.parse(stdout);
}

// Pull spreadsheet-shaped asset links out of markdown/text bodies. GitHub serves
// user uploads from a handful of hosts; we record every match and download only
// the ones that look like real fixtures (by extension).
const URL_RE = /https?:\/\/[^\s)"'<>\]]+/g;
function discoverAttachments(text) {
  if (!text) return [];
  const out = [];
  for (const url of text.match(URL_RE) ?? []) {
    const clean = url.replace(/[.,);]+$/, '');
    const lastSegment = clean.split('?')[0].split('#')[0].split('/').pop() ?? '';
    // Only treat a trailing `.foo` on the final path segment as an extension —
    // a bare URL like `.../pull/636` has no filename and no fixture.
    const match = lastSegment.match(/\.([a-z0-9]{1,5})$/i);
    const ext = match ? match[1].toLowerCase() : null;
    out.push({url: clean, ext, isFixture: ext !== null && SPREADSHEET_EXTENSIONS.has(ext)});
  }
  return out;
}

async function downloadFixture(url, destPath) {
  const res = await fetch(url, {redirect: 'follow'});
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  const length = Number(res.headers.get('content-length') ?? '0');
  if (length > MAX_ATTACHMENT_BYTES) {
    throw new Error(`Refusing ${length} bytes (> ${MAX_ATTACHMENT_BYTES} cap): ${url}`);
  }
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.byteLength > MAX_ATTACHMENT_BYTES) {
    throw new Error(`Refusing ${buf.byteLength} bytes (> ${MAX_ATTACHMENT_BYTES} cap): ${url}`);
  }
  await mkdir(dirname(destPath), {recursive: true});
  await writeFile(destPath, buf);
  return buf.byteLength;
}

function normalizeComment(c) {
  return {
    author: c.user?.login ?? null,
    authorAssociation: c.author_association ?? null,
    createdAt: c.created_at ?? null,
    reactions: c.reactions?.total_count ?? 0,
    body: c.body ?? '',
  };
}

async function main() {
  const {number, repo} = parseArgs(process.argv.slice(2));
  const padded = String(number).padStart(4, '0');

  const issue = await gh(`repos/${repo}/issues/${number}`);
  const isPR = Boolean(issue.pull_request);
  const comments = issue.comments > 0 ? await gh(`repos/${repo}/issues/${number}/comments`, {paginate: true}) : [];

  const attachmentSources = [issue.body, ...comments.map(c => c.body)].join('\n');
  const attachments = discoverAttachments(attachmentSources);

  // PRs carry their real value in the reproduction and the changed-file map, not
  // the diff itself (we are deleting the code the diff targets). Capture the shape.
  let pr;
  if (isPR) {
    const [detail, files] = await Promise.all([
      gh(`repos/${repo}/pulls/${number}`),
      gh(`repos/${repo}/pulls/${number}/files`, {paginate: true}),
    ]);
    pr = {
      merged: detail.merged ?? false,
      mergeable: detail.mergeable ?? null,
      baseRef: detail.base?.ref ?? null,
      additions: detail.additions ?? null,
      deletions: detail.deletions ?? null,
      changedFiles: files.map(f => ({path: f.filename, status: f.status, additions: f.additions, deletions: f.deletions})),
    };
  }

  const record = {
    schema: 'ts-xlsx/backlog-item@1',
    repo,
    number,
    type: isPR ? 'pull_request' : 'issue',
    url: issue.html_url,
    title: issue.title,
    state: issue.state,
    author: issue.user?.login ?? null,
    createdAt: issue.created_at ?? null,
    closedAt: issue.closed_at ?? null,
    labels: (issue.labels ?? []).map(l => (typeof l === 'string' ? l : l.name)),
    reactions: issue.reactions?.total_count ?? 0,
    commentCount: issue.comments ?? 0,
    body: issue.body ?? '',
    comments: comments.map(normalizeComment),
    attachments,
    pr,
  };
  // Triage/disposition ({captured|superseded|out-of-scope}) lives in
  // docs/knowledge/BACKLOG.md, not here — re-harvesting must refresh the raw
  // thread without ever clobbering a triage decision.

  const issuesDir = resolve(BACKLOG_DIR, 'issues');
  await mkdir(issuesDir, {recursive: true});
  const recordPath = resolve(issuesDir, `${padded}.json`);
  await writeFile(recordPath, `${JSON.stringify(record, null, 2)}\n`);

  const fixtures = attachments.filter(a => a.isFixture);
  const downloaded = [];
  for (const [i, att] of fixtures.entries()) {
    const name = `${padded}-${i}.${att.ext}`;
    const dest = resolve(BACKLOG_DIR, 'attachments', padded, name);
    try {
      const bytes = await downloadFixture(att.url, dest);
      downloaded.push({name, bytes});
    } catch (err) {
      downloaded.push({name, error: String(err.message ?? err)});
    }
  }

  console.log(`Harvested ${repo}#${number} (${record.type}) -> ${recordPath}`);
  console.log(`  ${comments.length} comment(s), ${attachments.length} link(s), ${fixtures.length} fixture candidate(s)`);
  for (const d of downloaded) {
    console.log(d.error ? `  ! ${d.name}: ${d.error}` : `  + ${d.name} (${d.bytes} bytes)`);
  }
}

main().catch(err => {
  console.error(`harvest failed: ${err.message ?? err}`);
  process.exitCode = 1;
});
