// The shared shape of a regression-corpus case.
//
// The corpus is *implementation-blind*: a case asserts observable behavior through
// an adapter surface (`api`) that is intentionally untyped, so a case can never
// couple to one implementation — that decoupling is the whole contract of the
// corpus (see test/corpus/README.md). What IS typed here is the case scaffold
// itself (id, cluster, baseline, the behavior list) and the `assert` module, so a
// typo in the harness — a missing `behavior`, a misspelled `assert.strictEqual` —
// is a compile error rather than a silent no-op at runtime.

import type assert from 'node:assert/strict';

/** The strict `node:assert` surface handed to every behavior. */
export type Assert = typeof assert;

/**
 * The adapter under test. Untyped by design: the corpus asserts behavior, never a
 * concrete implementation's types.
 */
// biome-ignore lint/suspicious/noExplicitAny: implementation-blind by contract — see above.
export type CorpusApi = any;

/** One observable behavior and the outcome we expect the implementation to produce. */
export interface Behavior {
  /** Human-readable statement of the behavior, shown by the runner. */
  name: string;
  /**
   * The outcome we currently expect: `pass` is a green regression lock; `fail`
   * marks a tracked known-open capability the rewrite has yet to build.
   */
  baseline: 'pass' | 'fail';
  /** Exercise the behavior; throw (via `assert`) to fail it. */
  expect(api: CorpusApi, assert: Assert): void | Promise<void>;
}

/** A disposable trace of where a case came from; never used to identify it. */
export interface Provenance {
  source?: string;
  repo?: string;
  ref?: number | string;
  url?: string;
  [key: string]: unknown;
}

/** A single corpus case: a durable id/cluster and the behaviors it locks in. */
export interface Case {
  /** Stable identity of the case; never derived from an upstream number. */
  id: string;
  /** The behavioral cluster this case belongs to. */
  cluster: string;
  /** What the case pins down, in prose. */
  description: string;
  provenance: Provenance;
  behavior: Behavior[];
}
