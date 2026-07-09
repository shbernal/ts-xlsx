// Shared core for the harvest toolchain.
//
// The harvest is a *one-time* operation: we pull the entire upstream backlog
// into a local queue once, then drain that queue by distilling each thread into
// durable product (a corpus case and/or spec note) and deleting the raw record.
// We never re-harvest to pick up new upstream activity — see STRATEGY.md Phase 1
// and docs/knowledge/BACKLOG.md for the drain model.
//
// This module holds the pieces the CLI entry points share (fetch-issue,
// harvest-all, list-backlog, status) so the fetch logic lives in exactly one
// place. Auth comes from the `gh` CLI (already logged in); we shell out to
// `gh api` rather than embedding a token so credentials never touch this process.

import {execFile} from 'node:child_process';
import {access, mkdir, writeFile} from 'node:fs/promises';
import {dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {promisify} from 'node:util';

const execFileAsync = promisify(execFile);

const HERE = dirname(fileURLToPath(import.meta.url));

export const REPO_ROOT = resolve(HERE, '..', '..');
export const BACKLOG_DIR = resolve(REPO_ROOT, 'docs', 'knowledge', 'backlog');
export const ISSUES_DIR = resolve(BACKLOG_DIR, 'issues');
export const ATTACHMENTS_DIR = resolve(BACKLOG_DIR, 'attachments');
export const MANIFEST_PATH = resolve(BACKLOG_DIR, 'manifest.json');
export const DEFAULT_REPO = 'exceljs/exceljs';

// A spreadsheet library ingests hostile input; the harvester ingests hostile
// URLs. Cap attachment downloads so a malicious/huge asset can't exhaust disk.
const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
const SPREADSHEET_EXTENSIONS = new Set(['xlsx', 'xlsm', 'xlsb', 'xls', 'csv', 'zip', 'ods']);

export function padNumber(number) {
  return String(number).padStart(4, '0');
}

export function recordPathFor(number) {
  return resolve(ISSUES_DIR, `${padNumber(number)}.json`);
}

export function isValidRepo(repo) {
  return /^[\w.-]+\/[\w.-]+$/.test(repo);
}

export async function fileExists(path) {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}

export async function gh(path, {paginate = false} = {}) {
  const argv = ['api', path];
  if (paginate) argv.push('--paginate');
  const {stdout} = await execFileAsync('gh', argv, {maxBuffer: 128 * 1024 * 1024});
  if (paginate) {
    // `gh --paginate` concatenates JSON arrays as `][` at page seams; splice them.
    return JSON.parse(stdout.replace(/\]\s*\[/g, ','));
  }
  return JSON.parse(stdout);
}

// GitHub's issues endpoint returns PRs too (they carry a `pull_request` key), so
// one paginated call enumerates the whole open universe. Used once by
// list-backlog to snapshot the manifest.
export async function listOpenItems(repo = DEFAULT_REPO) {
  const raw = await gh(`repos/${repo}/issues?state=open&per_page=100`, {paginate: true});
  return raw.map(item => ({
    number: item.number,
    type: item.pull_request ? 'pull_request' : 'issue',
    title: item.title ?? '',
    labels: (item.labels ?? []).map(l => (typeof l === 'string' ? l : l.name)),
    reactions: item.reactions?.total_count ?? 0,
    comments: item.comments ?? 0,
    url: item.html_url,
  }));
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

// Harvest a single thread into a durable JSON record plus any spreadsheet-shaped
// attachments. This is the atom Phase 1 fans out across the whole backlog.
// Returns a summary; with `skipIfExists`, a thread already on disk is left
// untouched so a bulk fill is cheaply resumable.
export async function harvestOne({number, repo = DEFAULT_REPO, skipIfExists = false} = {}) {
  if (!Number.isInteger(number) || number <= 0) {
    throw new Error(`Expected a positive issue/PR number, got: ${number}`);
  }
  if (!isValidRepo(repo)) {
    throw new Error(`Invalid repo: ${repo}`);
  }

  const padded = padNumber(number);
  const recordPath = recordPathFor(number);
  if (skipIfExists && (await fileExists(recordPath))) {
    return {number, recordPath, skipped: true};
  }

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

  await mkdir(ISSUES_DIR, {recursive: true});
  await writeFile(recordPath, `${JSON.stringify(record, null, 2)}\n`);

  const fixtures = attachments.filter(a => a.isFixture);
  const downloaded = [];
  for (const [i, att] of fixtures.entries()) {
    const name = `${padded}-${i}.${att.ext}`;
    const dest = resolve(ATTACHMENTS_DIR, padded, name);
    try {
      const bytes = await downloadFixture(att.url, dest);
      downloaded.push({name, bytes});
    } catch (err) {
      downloaded.push({name, error: String(err.message ?? err)});
    }
  }

  return {
    number,
    recordPath,
    skipped: false,
    type: record.type,
    commentCount: comments.length,
    attachmentCount: attachments.length,
    fixtureCount: fixtures.length,
    downloaded,
  };
}
