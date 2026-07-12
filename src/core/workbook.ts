// The workbook: the top of the model and the entry point of the public API.
//
// It owns its worksheets and the document-level properties. Sheet identity follows
// Excel's rules — names are unique case-insensitively, bounded in length, and free
// of the characters Excel forbids — so an invalid book cannot be constructed in the
// first place, rather than failing only at write time.

import type {WorkbookImage} from './image.ts';
import {Worksheet, type WorksheetState} from './worksheet.ts';

/** Document-level metadata written to the package's core properties. */
export interface WorkbookProperties {
  creator?: string;
  lastModifiedBy?: string;
  created?: Date;
  modified?: Date;
}

export interface AddWorksheetOptions {
  readonly state?: WorksheetState['state'];
}

/** A picture registered on the workbook, ready to be anchored to a worksheet. */
export interface AddImageOptions {
  /** The image bytes. */
  readonly buffer: Uint8Array;
  /** The file kind — `"png"`, `"jpeg"`/`"jpg"`, `"gif"`, … — without a leading dot. */
  readonly extension: string;
}

const MAX_SHEET_NAME_LENGTH = 31;
// Excel rejects these in a sheet name, plus a leading/trailing apostrophe.
const INVALID_SHEET_NAME_CHARS = /[*?:\\/[\]]/;

export class Workbook {
  readonly properties: WorkbookProperties = {};

  readonly #worksheets: Worksheet[] = [];
  #nextSheetId = 1;

  // Media is shared workbook-wide: a worksheet anchors an image by its registry index, so one
  // picture used on several sheets is stored once.
  readonly #media: WorkbookImage[] = [];

  /** The worksheets in insertion order. */
  get worksheets(): readonly Worksheet[] {
    return this.#worksheets;
  }

  /**
   * Register a picture on the workbook and return its numeric id. Pass the id to
   * {@link Worksheet.addImage} to anchor the picture to a sheet; the same id may be anchored on any
   * number of sheets and positions, and the bytes are still stored only once.
   */
  addImage(options: AddImageOptions): number {
    this.#media.push({extension: options.extension.toLowerCase(), data: options.buffer});
    return this.#media.length - 1;
  }

  /** The registered images, indexed by the id {@link addImage} returned. */
  get media(): readonly WorkbookImage[] {
    return this.#media;
  }

  /** Look up a registered image by its id, or `undefined` if no image carries that id. */
  getImage(id: number): WorkbookImage | undefined {
    return this.#media[id];
  }

  /**
   * Create a worksheet and append it to the workbook.
   *
   * @throws {Error} if the name is empty, too long, contains a forbidden character,
   *   or collides (case-insensitively) with an existing sheet.
   */
  addWorksheet(name: string, options: AddWorksheetOptions = {}): Worksheet {
    this.#assertValidSheetName(name);
    const sheet = new Worksheet(name, this.#nextSheetId++, options.state ?? 'visible');
    this.#worksheets.push(sheet);
    return sheet;
  }

  /** Look up a worksheet by name (case-insensitive) or by numeric id. */
  getWorksheet(nameOrId: string | number): Worksheet | undefined {
    if (typeof nameOrId === 'number') {
      return this.#worksheets.find(sheet => sheet.id === nameOrId);
    }
    const target = nameOrId.toLowerCase();
    return this.#worksheets.find(sheet => sheet.name.toLowerCase() === target);
  }

  #assertValidSheetName(name: string): void {
    if (name.length === 0) {
      throw new Error('worksheet name cannot be empty');
    }
    if (name.length > MAX_SHEET_NAME_LENGTH) {
      throw new Error(`worksheet name "${name}" exceeds the ${MAX_SHEET_NAME_LENGTH}-character limit`);
    }
    if (INVALID_SHEET_NAME_CHARS.test(name)) {
      throw new Error(`worksheet name "${name}" contains a character Excel forbids (* ? : \\ / [ ])`);
    }
    if (name.startsWith("'") || name.endsWith("'")) {
      throw new Error(`worksheet name "${name}" cannot start or end with an apostrophe`);
    }
    if (this.getWorksheet(name) !== undefined) {
      throw new Error(`a worksheet named "${name}" already exists (names are case-insensitive)`);
    }
  }
}
