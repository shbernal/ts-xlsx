// Excel-oracle harness — probe emitter.
//
// Turns a probe's declarative `spec` into a real `.xlsx` on disk, so the COM driver (observe.ps1) has
// something for Excel Desktop to open. It is a deliberately narrow, strictly-typed re-expression of the
// corpus adapter's `buildFrom` — NOT a reuse of it: that builder is `any`-typed corpus-internal
// machinery wired to JSZip and a module loader we do not want to drag into a standalone probe tool.
// The oracle only ever probes the handful of cell shapes below, so a small typed builder is the honest
// surface. Extend it as new invariants need new vocabulary.

import {mkdirSync, writeFileSync} from 'node:fs';
import {dirname} from 'node:path';

import {Workbook} from '../../src/core/workbook.ts';
import {writeXlsx} from '../../src/io/xlsx/write.ts';

/** A cached formula result — the `<v>` Excel would have computed, carried so the package is complete. */
export type FormulaResult = number | string | boolean;

/** One cell of a probe sheet: a literal value, a master formula, or a shared-formula clone. */
export type ProbeCell =
  | {readonly ref: string; readonly value: string | number | boolean}
  | {readonly ref: string; readonly formula: string; readonly result?: FormulaResult}
  | {readonly ref: string; readonly sharedFormula: string; readonly result?: FormulaResult};

export interface ProbeSheet {
  readonly name: string;
  readonly cells: readonly ProbeCell[];
}

export interface ProbeSpec {
  readonly sheets: readonly ProbeSheet[];
}

/** Build the in-memory workbook a probe spec describes. Exported so run.ts emits without re-shelling. */
export function buildWorkbook(spec: ProbeSpec): Workbook {
  const workbook = new Workbook();
  for (const sheetSpec of spec.sheets) {
    const sheet = workbook.addWorksheet(sheetSpec.name);
    for (const c of sheetSpec.cells) {
      const cell = sheet.getCell(c.ref);
      if ('formula' in c) {
        cell.value = 'result' in c ? {formula: c.formula, result: c.result} : {formula: c.formula};
      } else if ('sharedFormula' in c) {
        cell.value =
          'result' in c
            ? {sharedFormula: c.sharedFormula, result: c.result}
            : {sharedFormula: c.sharedFormula};
      } else {
        cell.value = c.value;
      }
    }
  }
  return workbook;
}

/** Serialize a probe spec to a `.xlsx` at `outPath`, creating parent directories as needed. */
export function emitProbe(spec: ProbeSpec, outPath: string): void {
  const bytes = writeXlsx(buildWorkbook(spec));
  mkdirSync(dirname(outPath), {recursive: true});
  writeFileSync(outPath, bytes);
}
