# ADR 0043: Main publishes an opt-in next channel

**Status:** Accepted (2026-10-13)

## Context

Stable releases are deliberate GitHub releases, and ADR-0026 keeps `latest` behind that act. A
dependent project that needs an unreleased fix had no package version to install, so it had to use a
Git URL or wait for a release. Both sibling libraries already publish prereleases under npm's `next`
dist-tag.

The repository also retained the inherited `master` branch name after its independent rewrite. The
other projects use `main`, as do new projects in this workspace.

## Decision

The default branch is `main`. Workflows, generated repository links, Pages deployment policy and
open pull requests follow the renamed branch.

Every push to `main` that can change the package publishes `@shbernal/ts-xlsx` under the `next`
dist-tag. The workflow derives a unique version shaped as `<base>-next.<run>.g<sha>`. A stable
`3.2.1` manifest therefore produces a version such as `3.2.2-next.42.g1a2b3c4`. The `g` keeps an
all-numeric abbreviated hash from becoming an invalid SemVer numeric identifier.

The workflow changes the version only in its checkout and runs the same `prepublishOnly` gate as a
stable release. It passes `--tag next` explicitly, so it cannot move `latest`. The `npm-publish`
environment admits `main` for this job and retains its `v*` tag rule for stable releases. npm trusted
publishing still restricts the package identity to this repository, workflow file and environment.

## Consequences

- A dependent project can install `@shbernal/ts-xlsx@next` and receive the newest published
  prerelease.
- Ordinary installs and ranges continue to resolve `latest`; prereleases are opt-in.
- A package-affecting push spends a new immutable npm version after the release gate passes.
- A merge that changes only tests, decision records or the website does not publish byte-identical
  package contents under another version.
