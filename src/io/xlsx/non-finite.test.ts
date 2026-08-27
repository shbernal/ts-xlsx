// A number that cannot be spelled in OOXML is refused at the moment it would become bytes.
//
// Every attribute exercised here is `xsd:double` or `xsd:unsignedInt` in ECMA-376, and each sits on
// a plain data shape (`PageSetup`, `PageBreak`, `Color`, `ConditionalFormattingRule`, `Extent`,
// `WorkbookView`, row/column properties) that the model stores verbatim with no accessor to validate
// through. So the write path is the one place the mistake can be caught, and it is caught per
// family: the helper being correct is not evidence a call site adopted it.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {AuthoringError} from '../../errors.ts';
import {writeXlsx} from './write.ts';

const UNWRITABLE = [Number.NaN, Infinity, -Infinity] as const;

// A workbook needs a sheet with content before any of this is reached.
function sheeted(): Workbook {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = 1;
  return workbook;
}

function refuses(mutate: (workbook: Workbook) => void): void {
  const workbook = sheeted();
  mutate(workbook);
  assert.throws(() => writeXlsx(workbook), AuthoringError);
}

test('<pageSetup> refuses a non-finite scaling attribute', () => {
  for (const value of UNWRITABLE) {
    refuses((wb) => {
      wb.getWorksheet('S')!.pageSetup.scale = value;
    });
    refuses((wb) => {
      wb.getWorksheet('S')!.pageSetup.fitToWidth = value;
    });
  }
});

test('<brk> refuses a non-finite row or column index', () => {
  refuses((wb) => {
    wb.getWorksheet('S')!.rowBreaks.push({id: Number.NaN});
  });
  refuses((wb) => {
    wb.getWorksheet('S')!.columnBreaks.push({id: 1, max: Infinity});
  });
});

test('a colour refuses a non-finite theme, tint or index', () => {
  for (const color of [{theme: Number.NaN}, {tint: Infinity}, {indexed: -Infinity}]) {
    refuses((wb) => {
      wb.getWorksheet('S')!.getCell('A1').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: color,
      };
    });
  }
});

test('<cfRule> refuses a non-finite rank or standard deviation', () => {
  refuses((wb) => {
    wb.getWorksheet('S')!.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [{type: 'top10', rank: Number.NaN, priority: 1}],
    });
  });
  refuses((wb) => {
    wb.getWorksheet('S')!.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [{type: 'aboveAverage', stdDev: Infinity, priority: 1}],
    });
  });
});

test('a picture refuses a non-finite extent or rotation', () => {
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  refuses((wb) => {
    const id = wb.addImage({buffer: png, extension: 'png'});
    // The EMU-level anchor, because the pixel-taking `addImage` would convert the bad value first
    // and the attribute is what is under test.
    wb.getWorksheet('S')!.addImageAnchor(id, {
      from: {col: 0, row: 0},
      ext: {cx: Number.NaN, cy: 100},
    });
  });
  refuses((wb) => {
    const id = wb.addImage({buffer: png, extension: 'png'});
    wb.getWorksheet('S')!.addImageAnchor(id, {
      from: {col: 0, row: 0},
      to: {col: 2, row: 2},
      rotation: Infinity,
    });
  });
});

test('an anchor grid point is refused on the same terms as the extent', () => {
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  refuses((wb) => {
    const id = wb.addImage({buffer: png, extension: 'png'});
    wb.getWorksheet('S')!.addImageAnchor(id, {
      from: {col: 0, row: 0, rowOff: Number.NaN},
      ext: {cx: 100, cy: 100},
    });
  });
});

test('<workbookView> refuses a non-finite window rect', () => {
  refuses((wb) => {
    wb.view.width = Number.NaN;
  });
});

test('an outline level is refused rather than dropped when it is not a number', () => {
  for (const value of UNWRITABLE) {
    refuses((wb) => {
      wb.getWorksheet('S')!.getColumn(2).outlineLevel = value;
    });
    refuses((wb) => {
      wb.getWorksheet('S')!.getRow(2).outlineLevel = value;
    });
  }
});
