// The field table behind `Worksheet.model`'s round-trip.
//
// Both directions used to be hand-written statement lists, one in the getter and one in the setter,
// each enumerating the same nineteen fields. Nothing but a comment asked the next editor to touch
// both, and a field exported but not imported loses data silently — the merge-loss failure the
// model contract exists to prevent. Here each field declares both directions in one place, and the
// registry is proved exhaustive over `keyof WorksheetModel` at compile time, so adding a field
// without wiring it is an error naming the field rather than a review catch.

import {cellToModel, copyCellContent} from './cell.ts';
import {cloneConditionalFormatting} from './conditional-formatting.ts';
import {overwrite, replaceContents} from './containers.ts';
import {cloneDataValidation} from './data-validation.ts';
import {INTERNAL} from './internal.ts';
import type {CellModel, Worksheet, WorksheetModel} from './worksheet.ts';

/** One field of a {@link WorksheetModel}, with both directions of its round-trip declared together. */
interface ModelFacet<K extends keyof WorksheetModel = keyof WorksheetModel> {
  readonly key: K;
  /** Produce the field's value, copied deeply enough that mutating it cannot reach back into the sheet. */
  readonly read: (sheet: Worksheet) => WorksheetModel[K];
  /**
   * Apply the field to a sheet whose content has already been reset. Takes the whole model rather
   * than the field so that a loop over the registry needs no correlation between `key` and the
   * field's type — {@link facet} does that projection once, where the key is still a single type.
   */
  readonly write: (sheet: Worksheet, model: WorksheetModel) => void;
}

function facet<K extends keyof WorksheetModel>(
  key: K,
  read: (sheet: Worksheet) => WorksheetModel[K],
  write: (sheet: Worksheet, value: WorksheetModel[K]) => void,
): ModelFacet<K> {
  return {key, read, write: (sheet, model) => write(sheet, model[key])};
}

/**
 * Every field of a {@link WorksheetModel}, in the order a model assignment applies them. Order is
 * load-bearing: cells are placed at their exact positions before any merge exists, so a covered
 * cell's value lands where the model says instead of being routed to a region master mid-load.
 */
export const WORKSHEET_MODEL_FACETS = [
  facet(
    'state',
    (sheet) => sheet.state,
    (sheet, value) => {
      sheet.state = value;
    },
  ),
  facet(
    'tabColor',
    (sheet) => sheet.tabColor,
    (sheet, value) => {
      sheet.tabColor = value;
    },
  ),
  facet(
    'properties',
    (sheet) => ({...sheet.properties}),
    (sheet, value) => overwrite(sheet.properties, value),
  ),
  facet(
    'outline',
    (sheet) => ({...sheet.outline}),
    (sheet, value) => overwrite(sheet.outline, value),
  ),
  facet(
    'pageSetup',
    (sheet) => ({...sheet.pageSetup}),
    (sheet, value) => overwrite(sheet.pageSetup, value),
  ),
  facet(
    'printOptions',
    (sheet) => ({...sheet.printOptions}),
    (sheet, value) => overwrite(sheet.printOptions, value),
  ),
  facet(
    'pageMargins',
    (sheet) => ({...sheet.pageMargins}),
    (sheet, value) => overwrite(sheet.pageMargins, value),
  ),
  facet(
    'headerFooter',
    (sheet) => ({...sheet.headerFooter}),
    (sheet, value) => overwrite(sheet.headerFooter, value),
  ),
  facet(
    'rowBreaks',
    (sheet) => sheet.rowBreaks.map((brk) => ({...brk})),
    (sheet, value) =>
      replaceContents(
        sheet.rowBreaks,
        value.map((brk) => ({...brk})),
      ),
  ),
  facet(
    'columnBreaks',
    (sheet) => sheet.columnBreaks.map((brk) => ({...brk})),
    (sheet, value) =>
      replaceContents(
        sheet.columnBreaks,
        value.map((brk) => ({...brk})),
      ),
  ),
  facet(
    'columns',
    (sheet) =>
      [...sheet.columns()].map(({index, properties}) => ({index, properties: {...properties}})),
    (sheet, value) => {
      for (const {index, properties} of value) Object.assign(sheet.getColumn(index), properties);
    },
  ),
  facet(
    'rows',
    (sheet) => {
      const rows: WorksheetModel['rows'] = [];
      // A row appears here only for its formatting; the cells it holds are the `cells` field's
      // business, and a row carrying nothing but cells has no properties to round-trip.
      for (const {number, properties} of sheet.rows()) {
        if (properties !== undefined) rows.push({number, properties: {...properties}});
      }
      return rows;
    },
    (sheet, value) => {
      for (const {number, properties} of value) Object.assign(sheet.getRow(number), properties);
    },
  ),
  facet(
    'cells',
    (sheet) => {
      const cells: CellModel[] = [];
      for (const row of sheet.rows()) for (const cell of row.cells) cells.push(cellToModel(cell));
      return cells;
    },
    (sheet, value) => {
      for (const cell of value) copyCellContent(cell, sheet[INTERNAL].cellAt(cell.row, cell.col));
    },
  ),
  facet(
    'merges',
    (sheet) => [...sheet.merges],
    (sheet, value) => {
      for (const range of value) sheet.mergeCells(range);
    },
  ),
  facet(
    'dataValidations',
    (sheet) =>
      sheet.dataValidations.map(({sqref, rule, extended}) => ({
        sqref,
        rule: cloneDataValidation(rule),
        ...(extended ? {extended: true} : {}),
      })),
    (sheet, value) => {
      for (const {sqref, rule, extended} of value) {
        sheet.addDataValidation(sqref, rule, extended ? {extended: true} : {});
      }
    },
  ),
  facet(
    'conditionalFormattings',
    (sheet) => sheet.conditionalFormattings.map(cloneConditionalFormatting),
    (sheet, value) => {
      for (const formatting of value) sheet.addConditionalFormatting(formatting);
    },
  ),
  facet(
    'tables',
    (sheet) => sheet.tables.map((table) => table.options),
    (sheet, value) => {
      for (const options of value) sheet.addTable(options);
    },
  ),
  facet(
    'autoFilter',
    (sheet) => sheet.autoFilter,
    // Through the public setter, which re-canonicalises the range and — on `undefined` — clears any
    // autofilter the destination held. That clearing is why the field is applied even when absent.
    (sheet, value) => {
      sheet.autoFilter = value;
    },
  ),
  facet(
    'protection',
    (sheet) => sheet.protection,
    // A loaded credential is already hashed, so it is reinstated verbatim rather than re-derived;
    // `protect` cannot express that, which is why this goes through the internal channel.
    (sheet, value) => {
      if (value === undefined) sheet.unprotect();
      else sheet[INTERNAL].restoreProtection(value);
    },
  ),
];

type AssertNever<T extends never> = T;

/**
 * Compile-time proof that {@link WORKSHEET_MODEL_FACETS} covers every {@link WorksheetModel} field.
 * A field added without a facet resolves this to that field's name, which does not satisfy `never`,
 * so the error names what is missing. This is the guarantee the registry exists to provide: the two
 * directions of the round-trip can no longer drift apart without the build saying so.
 */
export type EveryWorksheetModelFieldHasAFacet = AssertNever<
  Exclude<keyof WorksheetModel, (typeof WORKSHEET_MODEL_FACETS)[number]['key']>
>;
