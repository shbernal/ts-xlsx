// The streaming writer's zip container and the destinations its bytes go to. Entries are opened as the
// writer reaches them, which for a streamed sheet is its first committed row, long before the rest of
// the package exists (ADR 0046). Each output chunk is handed to the destinations the moment fflate
// produces it, and the writer asks `ready()` whether the caller's sink wants it to wait.

import {once} from 'node:events';
import type {PassThrough, Writable} from 'node:stream';

import {strToU8, Zip, ZipDeflate} from 'fflate';

import {concat} from '../../bytes.ts';
import {FIXED_ENTRY_MTIME} from '../opc/zip-mtime.ts';

// How much uncompressed data is pushed into the deflater between two looks at the sink. It bounds what
// one push can emit, so it is also roughly the most compressed output a slow sink is handed before the
// writer next asks whether it may continue. A streamed sheet's rows are batched to the same size, so a
// row is never pushed alone: the deflater is called once per batch, not once per row.
export const PUSH_SLICE = 64 * 1024;

/**
 * One zip archive written incrementally. Not reusable: {@link finish} ends it.
 *
 * A sink's failure is recorded rather than thrown. From then on nothing more is written to it, and
 * {@link failure} is what the workbook's `commit()` rejects with.
 */
export class PackageStream {
  readonly #sink: Writable | undefined;
  #collected: Uint8Array[] | null;
  #owned: PassThrough | undefined;
  #zip: Zip | undefined;
  #zipError: Error | undefined;
  #congested = false;
  #waiting: Promise<void> | undefined;
  #failure: Error | undefined;
  // The sink's terminal state, tracked from construction: output starts at the first committed row,
  // so a sink that fails to open (a bad filename errors on a later tick) has to be noticed whenever it
  // happens rather than only once the workbook commits.
  readonly #sinkSettled: Promise<void> | undefined;

  /**
   * `sink` is the caller's destination, if any. Without one every chunk is kept, so the archive can be
   * handed back whole; with one nothing is, unless {@link attach} asks for a copy before output starts.
   */
  constructor(sink: Writable | undefined) {
    this.#sink = sink;
    this.#collected = sink === undefined ? [] : null;
    if (this.#sink !== undefined) {
      const sink = this.#sink;
      this.#sinkSettled = new Promise((resolve, reject) => {
        sink.once('finish', resolve);
        sink.once('error', reject);
      });
      // Claimed here so a sink that errors while nobody is awaiting it is not an unhandled rejection;
      // the error is kept and reported by the workbook's `commit()`.
      this.#sinkSettled.catch((error: unknown) => {
        this.#failure ??= error instanceof Error ? error : new Error(String(error));
      });
    }
  }

  /** Whether any byte has been produced yet. */
  get started(): boolean {
    return this.#zip !== undefined;
  }

  /** The sink's error, once it has failed. */
  get failure(): Error | undefined {
    return this.#failure ?? this.#zipError;
  }

  /**
   * Also write every chunk into `owned`, starting with the ones already produced, and keep a copy of
   * the archive to hand back: a `PassThrough` nobody drained would otherwise be its only copy. With a
   * sink, the caller checks that output has not started, since those chunks were never kept.
   */
  attach(owned: PassThrough): void {
    this.#owned = owned;
    this.#collected ??= [];
    for (const chunk of this.#collected) owned.write(chunk);
  }

  /** Open a new entry. Its bytes are emitted in the order entries were opened. */
  open(name: string): PackageEntry {
    const entry = new ZipDeflate(name, {level: 6});
    // `mtime` is a field on the entry rather than a compression option (`ZipDeflate` forwards its
    // options to the deflater alone) and the header is written when the entry is added, so it has to
    // be set before that call rather than before the first push.
    entry.mtime = FIXED_ENTRY_MTIME;
    this.#archive().add(entry);
    return new PackageEntry(entry, this);
  }

  /**
   * Resolve once the sink can take more: at once when it has not asked to wait, otherwise on its
   * `drain`. Never rejects; a sink that errors instead resolves it, and the error is {@link failure}.
   * The writer-owned stream is not waited on: a caller may leave it unread until the workbook commits.
   */
  ready(): Promise<void> {
    if (!this.#congested || this.#failure !== undefined) return Promise.resolve();
    this.#waiting ??= this.#drained().finally(() => {
      this.#congested = false;
      this.#waiting = undefined;
    });
    return this.#waiting;
  }

  /** Write a whole part, in slices, waiting on the sink between them. */
  async writePart(name: string, data: Uint8Array): Promise<void> {
    const entry = this.open(name);
    let offset = 0;
    do {
      const end = Math.min(offset + PUSH_SLICE, data.length);
      entry.push(data.subarray(offset, end), end === data.length);
      offset = end;
      await this.ready();
    } while (offset < data.length);
  }

  /**
   * End the archive and the destinations, resolving with the whole archive when it was collected.
   *
   * @throws the sink's error, or the zip layer's, if either failed.
   */
  async finish(): Promise<Uint8Array | undefined> {
    this.#archive().end();
    this.#throwIfFailed();
    this.#owned?.end();
    this.#sink?.end();
    await this.#sinkSettled;
    this.#throwIfFailed();
    return this.#collected === null ? undefined : concat(this.#collected);
  }

  /** Destroy every destination with `failure`: the bytes written so far are not a whole package. */
  abort(failure: Error): void {
    // `destroy(err)`, never `end()`: ending says the package is complete, and a consumer handed a
    // truncated archive that ended cleanly will read it as whole. Destroying propagates through `pipe`
    // as the destination's `error`, which is what a `pipeline()` caller already handles.
    this.#owned?.destroy(failure);
    this.#sink?.destroy(failure);
  }

  #throwIfFailed(): void {
    const failure = this.failure;
    if (failure !== undefined) throw failure;
  }

  #archive(): Zip {
    this.#zip ??= new Zip((error, chunk) => {
      if (error) {
        this.#zipError ??= error;
        return;
      }
      this.#emit(chunk);
    });
    return this.#zip;
  }

  #emit(chunk: Uint8Array): void {
    this.#collected?.push(chunk);
    this.#owned?.write(chunk);
    const sink = this.#sink;
    if (sink === undefined || this.#failure !== undefined) return;
    if (!sink.write(chunk)) this.#congested = true;
  }

  // Wait until a sink that refused more input has drained, or until it fails: `#sinkSettled` rejects
  // on `error`, which a sink that errored before the `drain` would otherwise leave this waiting on.
  async #drained(): Promise<void> {
    const sink = this.#sink;
    const settled = this.#sinkSettled;
    if (sink === undefined || settled === undefined) return;
    const abort = new AbortController();
    try {
      await Promise.race([once(sink, 'drain', {signal: abort.signal}), settled]);
    } catch {
      // The failure is recorded by the handler on `#sinkSettled`; resolving is the contract.
    } finally {
      abort.abort();
    }
  }
}

/**
 * One open entry of a {@link PackageStream}. Text written to it is batched and deflated in
 * {@link PUSH_SLICE} pieces, so the memory it holds is one batch, whatever the entry's length.
 */
export class PackageEntry {
  readonly #entry: ZipDeflate;
  readonly #stream: PackageStream;
  #pending = '';
  #closed = false;

  constructor(entry: ZipDeflate, stream: PackageStream) {
    this.#entry = entry;
    this.#stream = stream;
  }

  /** Append text, deflating once a batch is full. */
  write(text: string): void {
    this.#pending += text;
    if (this.#pending.length >= PUSH_SLICE) this.#flush(false);
  }

  /** Raw bytes, for a part that is not text. `final` ends the entry. */
  push(bytes: Uint8Array, final: boolean): void {
    this.#entry.push(bytes, final);
    if (final) this.#closed = true;
  }

  /** Write what is left and end the entry. */
  close(): void {
    if (this.#closed) return;
    this.#flush(true);
  }

  /** {@link PackageStream.ready}. */
  ready(): Promise<void> {
    return this.#stream.ready();
  }

  #flush(final: boolean): void {
    const text = this.#pending;
    this.#pending = '';
    this.#entry.push(strToU8(text), final);
    if (final) this.#closed = true;
  }
}
