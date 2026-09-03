// The workbook's picture registry: the images themselves, and the content index that keeps
// re-registering an identical one off a byte-by-byte walk of everything already held.
//
// A slice of `Workbook` in the sense "the two model classes delegate their state" describes, and it
// reaches nothing outside itself. `exportImages`/`importImages` deliberately stay on `Workbook`,
// because they touch five things on `Worksheet` and moving them would relocate that coupling rather
// than remove it; they read through this. What is here is the registry and the one invariant it has
// to keep, which is that the *first* id registered under a content key wins.

import {AuthoringError, quoted} from '../errors.ts';
import {imageContentKey, normalizeImageExtension, type WorkbookImage} from './image.ts';

export class WorkbookMedia {
  // Shared workbook-wide: a worksheet anchors an image by its registry index, so one picture used on
  // several sheets is stored once.
  readonly #images: WorkbookImage[] = [];
  // Content key to the id of the first picture registered under it. Built on the first import that
  // needs it and kept in step by `register` thereafter, so a workbook nobody merges into never
  // digests a byte.
  #byContent: Map<string, number> | undefined;

  /** The registered images, indexed by the id {@link register} returned. Live, not a copy. */
  get all(): readonly WorkbookImage[] {
    return this.#images;
  }

  /** Look up a registered image by its id, or `undefined` if no image carries that id. */
  get(id: number): WorkbookImage | undefined {
    return this.#images[id];
  }

  /**
   * {@link get} for a caller that has no answer for an absent id.
   *
   * An unregistered id is what a sheet belonging to *another* workbook looks like from here, which
   * is why the message names the sheet rather than the id alone. Emitting a package with a drawing
   * pointing at media nobody registered is the silently-broken-image failure this refuses to start.
   *
   * @throws {AuthoringError} if no image carries that id.
   */
  require(id: number, forSheetName: string): WorkbookImage {
    const image = this.#images[id];
    if (image === undefined) {
      throw new AuthoringError(
        `worksheet ${quoted(forSheetName)} shows image id ${id}, which is not registered on this ` +
          "workbook: a sheet's images can only be exported by the workbook that holds them",
      );
    }
    return image;
  }

  /** Store a picture's bytes and return the id that names them. */
  register(extension: string | undefined, buffer: Uint8Array): number {
    const image: WorkbookImage = {
      extension: normalizeImageExtension(extension, buffer),
      data: buffer,
    };
    this.#images.push(image);
    const id = this.#images.length - 1;
    // Kept in step only once an import has built it. The first id wins on a repeat, which is the
    // answer the scan this replaced gave.
    const index = this.#byContent;
    if (index !== undefined) {
      const key = imageContentKey(image);
      if (!index.has(key)) index.set(key, id);
    }
    return id;
  }

  /**
   * Register a picture arriving from elsewhere, re-using an identical one already held.
   *
   * The extension is re-normalised rather than trusted: a hand-built {@link WorkbookImage} may carry
   * `".PNG"` where the registry holds `"png"`, and two spellings of one kind must not read as two
   * pictures.
   *
   * Through the content index rather than a scan. Comparing byte-by-byte against every held picture
   * made importing n distinct images cost n squared byte comparisons: fifty 1 MB pictures carried
   * between workbooks compared about 2.5 GB. The *rule* is unchanged, and `imageContentKey` states
   * why identity here is content and never object identity. The index is built in reverse so the
   * FIRST id registered under a key wins, which is the answer the scan gave.
   */
  registerExisting(image: WorkbookImage): number {
    const candidate: WorkbookImage = {
      extension: normalizeImageExtension(image.extension, image.data),
      data: image.data,
    };
    this.#byContent ??= new Map(
      this.#images.map((held, id): [string, number] => [imageContentKey(held), id]).reverse(),
    );
    return (
      this.#byContent.get(imageContentKey(candidate)) ??
      this.register(candidate.extension, candidate.data)
    );
  }
}
