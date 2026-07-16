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

/**
 * A named reference in the workbook — the entries Excel surfaces in its Name Manager. A name maps
 * an identifier to a formula (`refersTo`), most often a cell range like `Sheet1!$A$1:$B$2` but
 * possibly any formula. A name is global to the workbook unless it names a sheet in {@link scope},
 * which restricts it to that sheet and lets another sheet reuse the same name independently.
 */
export interface DefinedName {
  /** The name as typed in a formula, e.g. `"TaxRate"`. Built-in names carry an `_xlnm.` prefix. */
  readonly name: string;
  /** The formula the name resolves to, e.g. `"Sheet1!$A$1:$B$2"`. */
  readonly refersTo: string;
  /** The sheet the name is scoped to; omit for a workbook-global name. */
  readonly scope?: string;
  /** A human note shown beside the name in Excel's Name Manager. */
  readonly comment?: string;
  /** Hide the name from the Name Manager UI without removing it. */
  readonly hidden?: boolean;
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

  /**
   * Ask consuming spreadsheet apps to recalculate every formula when the file is opened, rather than
   * trusting the cached results stored with each formula cell. Set this when the producer cannot
   * compute formula results itself — the OOXML `fullCalcOnLoad` flag. Off by default, so a workbook
   * whose cached results are authoritative stays unmarked.
   */
  fullCalcOnLoad = false;

  readonly #worksheets: Worksheet[] = [];
  #nextSheetId = 1;

  // Media is shared workbook-wide: a worksheet anchors an image by its registry index, so one
  // picture used on several sheets is stored once.
  readonly #media: WorkbookImage[] = [];

  readonly #definedNames: DefinedName[] = [];

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

  /** The workbook's defined names, in the order they were registered. */
  get definedNames(): readonly DefinedName[] {
    return this.#definedNames;
  }

  /**
   * Register a defined name on the workbook.
   *
   * @throws {Error} if the name is empty, or if a {@link DefinedName.scope} is given that names no
   *   existing worksheet — a scoped name must target a sheet that is already part of the workbook.
   */
  defineName(definedName: DefinedName): void {
    if (definedName.name.length === 0) {
      throw new Error('a defined name cannot be empty');
    }
    if (definedName.scope !== undefined && this.getWorksheet(definedName.scope) === undefined) {
      throw new Error(`defined name "${definedName.name}" is scoped to unknown worksheet "${definedName.scope}"`);
    }
    this.#definedNames.push(definedName);
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
