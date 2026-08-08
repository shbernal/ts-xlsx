// The shared shape of a regression-corpus case.
//
// A case asserts observable behavior through the adapter surface (`api`), and it stays blind to how
// the library is built — it can reach `src` only through a named capability, which is the decoupling
// that let the corpus outlive the rewrite (see test/corpus/README.md). That surface is now *typed*:
// blindness to an implementation's internals is the contract, blindness to the capability list never
// was, and pretending not to know it cost every case its type-checking.

import type assert from 'node:assert/strict';
import type {CorpusApi} from './adapters/ts-xlsx.ts';

/** The strict `node:assert` surface handed to every behavior. */
export type Assert = typeof assert;

export type {CorpusApi};

/** One observable behavior the corpus locks in. */
export interface Behavior {
  /** Human-readable statement of the behavior, shown by the runner. */
  name: string;
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
