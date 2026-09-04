// Running a PowerShell script and getting its output back, bounded on both axes.
//
// The two Windows-only harnesses (`excel-oracle`, `vba-compiler`) each carried a byte-identical copy
// of this, forty-four lines apiece: the spawn flags, the output cap, the settled latch, the SIGKILL
// timeout. Their timeouts had already drifted, 120 s against 180 s, which is the only difference
// that was ever meant to be one and so is the one thing this takes as an argument. Neither harness
// runs in CI, so a bug in one copy is invisible until somebody is mid-investigation, which is the
// worst moment for a tool to be the thing that is wrong.
//
// Both bounds are here because a host process is being asked to drive a GUI application over COM.
// It can hang, and it can produce output without end; neither is hypothetical, and a harness that
// does either takes the investigation with it.

import {spawn} from 'node:child_process';

/** 8 MB of combined stdout and stderr. An observation larger than this is a runaway, not a result. */
const MAX_OUTPUT_BYTES = 8 * 1024 * 1024;

export interface PwshResult {
  /** The exit code, or `null` when the run was killed (timed out, or capped) or never started. */
  readonly code: number | null;
  readonly stdout: string;
  readonly stderr: string;
  /** Set when the spawn itself failed, which is what a host with no `pwsh` looks like. */
  readonly spawnError?: Error;
}

/**
 * Run `pwsh` with the given arguments, resolving rather than rejecting.
 *
 * Every failure mode arrives as a `PwshResult`: a caller of a tool like this wants to report what
 * happened, and a rejection would make each of the four outcomes a separate `catch`. `-NoProfile`
 * and `-NonInteractive` keep the host's own configuration and any prompt out of the run.
 */
export function runPwsh(args: readonly string[], timeoutMs: number): Promise<PwshResult> {
  return new Promise<PwshResult>((resolve) => {
    const child = spawn('pwsh', ['-NoProfile', '-NonInteractive', ...args], {
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    const stdout: Buffer[] = [];
    const stderr: Buffer[] = [];
    let bytes = 0;
    // The timeout, the cap, the spawn error and the close can all fire, and two of them fire
    // *because* of the third: killing the child on a timeout also closes it.
    let settled = false;
    const done = (result: PwshResult): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(result);
    };
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      done({code: null, stdout: Buffer.concat(stdout).toString('utf8'), stderr: 'timed out'});
    }, timeoutMs);
    const collect = (target: Buffer[]) => (chunk: Buffer) => {
      bytes += chunk.length;
      if (bytes > MAX_OUTPUT_BYTES) {
        child.kill('SIGKILL');
        done({
          code: null,
          stdout: Buffer.concat(stdout).toString('utf8'),
          stderr: 'output too large',
        });
        return;
      }
      target.push(chunk);
    };
    child.stdout.on('data', collect(stdout));
    child.stderr.on('data', collect(stderr));
    child.on('error', (spawnError) => done({code: null, stdout: '', stderr: '', spawnError}));
    child.on('close', (code) =>
      done({
        code,
        stdout: Buffer.concat(stdout).toString('utf8'),
        stderr: Buffer.concat(stderr).toString('utf8'),
      }),
    );
  });
}

/** A harness's refusal to continue, prefixed with its own name. */
export function failWith(tool: string): (message: string) => never {
  return (message: string): never => {
    process.stderr.write(`${tool}: ${message}\n`);
    process.exit(1);
  };
}
