# ADR 0026 — Releasing is a GitHub release, and npm follows with no credential

**Status:** Accepted (2026-07-29) · release slice, immediately after 1.0.0

## Context

ADR-0015 settled *what* a version is called and left publishing as "a separate,
deliberate, human-triggered action". 1.0.0 was then published by hand: `npm publish`
from a maintainer's terminal, with a one-time password typed at the prompt. That works
exactly once per person who has the password, and it produces a package the registry
cannot vouch for — npm serves the bytes and can say nothing about which commit or which
tree produced them. It also means every release is a chance to publish a dirty tree, a
wrong version, or a build nobody gated.

The obvious fix — a workflow holding an npm automation token — trades one problem for
another. A long-lived credential in repository settings is exfiltratable by any workflow
change that can read secrets, and it is invisible: nothing in a diff shows who can publish.

## Decision

1. **Publishing a GitHub release publishes to npm.** `.github/workflows/publish.yml`
   triggers on `release: published`. Cutting a release is therefore the whole ritual;
   there is no second command to remember and no terminal that must be someone's.

2. **No npm credential exists in this repository.** Authentication is npm **trusted
   publishing**: the job proves its identity with a GitHub OIDC token (`id-token: write`),
   and npm exchanges it for a short-lived one *only* when the repository, workflow
   filename and environment match the trusted publisher configured on the package. The
   same exchange emits a provenance attestation, so the registry can state which commit
   and which run built what it serves. There is nothing to leak and nothing to rotate.

3. **The publish job runs in the `npm-publish` environment, and that environment is
   code.** `.github/workflows/environment.yml` provisions it — a required reviewer, and
   deployments only from `v*` **tags** — so the gate guarding a publish is reviewable in a
   diff instead of being a settings page nobody reads. It needs a PAT with Administration
   rights, because `administration` is not among the permissions a workflow may request
   for `GITHUB_TOKEN`; the environments API is deliberately out of that token's reach.

4. **The workflow enumerates no gates.** `npm publish` runs `prepublishOnly` — build,
   `verify --full`, `smoke:dist`, the size budgets — which is the single definition of
   "everything" that ADR-0022 established and corpus.yml already refuses to duplicate.

5. **Three refusals are wired in**, because a release is the one action that cannot be
   taken back: a tag whose name disagrees with the version its own commit declares fails
   the run; a version already on the registry is a no-op rather than a red X, so re-running
   a workflow is safe; and a pre-release tag takes the `next` dist-tag, never `latest`.

## Consequences

- **Positive:** a release is one action with one approval; the published package carries
  provenance; no credential exists to steal; and who may publish is a file, not a memory.
- **Negative / deferred:** the trusted publisher must be configured once on npmjs.com
  before CI can publish at all, and it pins the workflow *filename* — renaming
  `publish.yml` breaks publishing until npm is updated to match. `environment.yml` needs a
  PAT, which is the one long-lived credential left; it can only reconfigure an environment,
  never publish.
- **Verified, not assumed** — and the verification corrected the design twice, once against
  a mistake of ours.

  `npm publish --dry-run` is not local. It builds the tarball and then asks the registry,
  which refuses a version it already serves, so a rehearsal reaches the publish step only
  for a version that is not out yet.

  `setup-node`'s `registry-url` is **required**. It sets `NODE_AUTH_TOKEN` to the literal
  placeholder `XXXXX-XXXXX-XXXXX-XXXXX`, from which we inferred that npm would send that
  placeholder instead of exchanging an OIDC token, and dropped the option. That inference
  was wrong, and 1.0.1 failed `ENEEDAUTH` proving it: `registry-url` writes the .npmrc
  entry naming the registry this publish authenticates against, and without it npm does not
  begin the exchange at all. The placeholder is genuinely unused — that variable is read
  when *installing* private dependencies, not when publishing. The lesson generalises: a
  plausible reading of an observed value is not evidence, and the release path is one place
  where the difference is a burned version number.
- **Revisit when:** npm changes the trusted-publishing contract, or a second package ships
  from this repository (the environment and the publisher config are both single-package
  shaped today).
