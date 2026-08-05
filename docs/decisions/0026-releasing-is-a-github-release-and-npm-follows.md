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
   code.** `.github/workflows/environment.yml` provisions it — deployments only from `v*`
   **tags**, and no required reviewer — so the gate guarding a publish is reviewable in a
   diff instead of being a settings page nobody reads. It needs a PAT with Administration
   rights, because `administration` is not among the permissions a workflow may request
   for `GITHUB_TOKEN`; the environments API is deliberately out of that token's reach.

   The environment carried a required reviewer through 1.0.3 and no longer does
   (**amended 2026-08-05**). On a single-maintainer project it was never a second pair of
   eyes: the only login that could approve was the one that had just cut the release, so it
   asked the releaser to confirm a decision they had made one command earlier. Consent that
   only the decider can give is a delay, not a control. What it genuinely stopped was an
   *accidental* publish — see the consequences below for what now carries that.

4. **The workflow enumerates no gates.** `npm publish` runs `prepublishOnly` — build,
   `verify --full`, `smoke:dist`, the size budgets — which is the single definition of
   "everything" that ADR-0022 established and corpus.yml already refuses to duplicate.

5. **Three refusals are wired in**, because a release is the one action that cannot be
   taken back: a tag whose name disagrees with the version its own commit declares fails
   the run; a version already on the registry is a no-op rather than a red X, so re-running
   a workflow is safe; and a pre-release tag takes the `next` dist-tag, never `latest`.

## Consequences

- **Positive:** a release is one action and nothing else; the published package carries
  provenance; no credential exists to steal; and who may publish is a file, not a memory.
- **The blast radius of a `v*` tag is now the whole gate.** With no reviewer, pushing a tag
  and publishing a release is the entire path to the registry, so anything holding push
  access — a leaked token, an agent given more rope than intended, a mistyped tag on the
  wrong commit — publishes unattended, and a version number cannot be taken back. Two
  things still stand between a mistake and npm: the tag-name/version agreement check, which
  refuses a tag whose commit declares a different version, and `prepublishOnly`, which
  refuses a tree that fails its gates. Neither catches a *correct* publish of something
  nobody meant to release. Rehearse (`dry_run`) when a release is in any way unusual;
  that habit is the replacement for the click, and unlike the click it can actually fail.
- **Negative / deferred:** the trusted publisher must be configured once on npmjs.com
  before CI can publish at all, and it pins three names that live outside this repository's
  review: the repository, the workflow *filename*, and the *environment*. Renaming
  `publish.yml` or the `npm-publish` environment breaks publishing until npm is updated to
  match, and npm reports any of the three disagreeing identically.

  **`ENV_ADMIN_TOKEN` is absent as of 2026-08-05, so `environment.yml` cannot run.** Removing
  the reviewer was therefore applied by hand — a `PUT` to the environments API with the same
  payload the workflow builds, read back and confirmed (`reviewers: none`, `deploys from: tag
  v*`). The file and the live environment agree today, but they agree by coincidence rather
  than by construction, which is the one property this ADR bought. Until the secret is
  restored, treat `environment.yml` as documentation: it describes the environment and can no
  longer enforce it, and a settings-page edit would drift silently the way it did before.
  `environment.yml` needs a PAT, which is the one long-lived credential left; it can only
  reconfigure an environment,
  never publish.
- **Verified, not assumed** — and the verification corrected the design twice and the
  npm-side configuration once, each against a mistake of ours.

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

  The publisher's **environment** field was then wrong on npm, and being wrong looked like
  nothing. It named `environment.yml` — the workflow that *provisions* the deployment
  environment — where it had to name the environment itself, `npm-publish`. npm answers a
  rejected identity with a 404, which reads as "no such package" and says nothing about
  which of the three claims failed to match, so 1.0.2 failed against a workflow that was
  already correct and the two commits after its tag sharpened the error message rather than
  the cause. The general rule that falls out: a publish that 404s on a package that
  demonstrably exists is an identity mismatch, and the thing to read is the trusted
  publisher on npmjs.com, not this repository.

  The rehearsal had been unable to catch any of this, which is the part worth keeping.
  `npm publish --dry-run` treats a rejected identity as a warning and exits `0` — so the
  rehearsal passed *because* it was a rehearsal, and the misconfiguration reached a real
  release twice. A check that cannot fail on the thing it is checking is decoration, and it
  reads as reassurance, which is worse than having no check at all. The job now reads the
  verbose log and fails when the exchange was rejected, or when npm never attempted one.
  Both markers were taken from the failing run's own output and re-run against it, rather
  than written from what the log was assumed to say — the same discipline the two corrections
  above had to be learned through.
- **Revisit when:** npm changes the trusted-publishing contract, or a second package ships
  from this repository (the environment and the publisher config are both single-package
  shaped today).
