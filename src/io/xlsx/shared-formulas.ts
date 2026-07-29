// Shared-formula group planning: assigning each master/clone group a sheet-unique index and the
// master's `ref` range before the row loop serialises a single cell.

import {encodeAddress} from '../../core/address.ts';
import type {Cell} from '../../core/cell.ts';
import {isFormulaValue, isSharedFormulaValue} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';

// A cell's role in an OOXML shared-formula group. A master carries the source formula plus the `ref`
// range the group spans; a clone (no `ref`) references the master's formula by the shared index `si`.
export interface SharedFormulaRole {
  readonly si: number;
  readonly ref?: string;
}

// Plan a sheet's shared-formula groups: every clone cell (a {@link SharedFormulaValue}) names its
// master by address, so group the clones by master, assign each group a sheet-unique `si`, and record
// the `ref` range (master through the furthest clone) on the master. Excel requires the master to sit
// at the top-left of that range, so a clone above or left of its master — or a master with no formula
// (an orphan) — is rejected here, named, rather than emitted as a package Excel repairs on open.
export function planSharedFormulas(sheet: Worksheet): Map<string, SharedFormulaRole> {
  const groups = new Map<string, Cell[]>();
  for (const {cells} of sheet.rows()) {
    for (const cell of cells) {
      if (isSharedFormulaValue(cell.value)) {
        const clones = groups.get(cell.value.sharedFormula);
        if (clones !== undefined) clones.push(cell);
        else groups.set(cell.value.sharedFormula, [cell]);
      }
    }
  }

  const roles = new Map<string, SharedFormulaRole>();
  let si = 0;
  for (const [masterAddress, clones] of groups) {
    const master = sheet.getCell(masterAddress);
    if (!isFormulaValue(master.value)) {
      const offender = clones[0] as Cell;
      throw new AuthoringError(
        `shared-formula clone ${offender.address} names master ${masterAddress}, which holds no formula`,
      );
    }
    let maxCol = master.col;
    let maxRow = master.row;
    for (const clone of clones) {
      if (clone.col < master.col || clone.row < master.row) {
        throw new AuthoringError(
          `shared-formula master ${masterAddress} must sit above and/or left of clone ${clone.address}`,
        );
      }
      if (clone.col > maxCol) maxCol = clone.col;
      if (clone.row > maxRow) maxRow = clone.row;
    }
    roles.set(masterAddress, {
      si,
      ref: `${encodeAddress(master.col, master.row)}:${encodeAddress(maxCol, maxRow)}`,
    });
    for (const clone of clones) roles.set(clone.address, {si});
    si += 1;
  }
  return roles;
}
