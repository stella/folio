import { describe, expect, test } from "bun:test";

import {
  resolveTableCellConditionalStyle,
  type TableConditionalStyles,
} from "./tableConditionalFormatting";

const styles = {
  wholeTable: { rPr: { color: { rgb: "111111" } } },
  nwCell: { rPr: { bold: true } },
  lastCol: { rPr: { strike: true } },
  band1Vert: { rPr: { italic: true } },
  band2Vert: { rPr: { smallCaps: true } },
} satisfies TableConditionalStyles;

const resolve = (
  overrides: Partial<Parameters<typeof resolveTableCellConditionalStyle>[0]> = {},
) =>
  resolveTableCellConditionalStyle({
    styles,
    look: { firstRow: true, firstColumn: true, lastColumn: true },
    rowIndex: 0,
    totalRows: 2,
    columnIndex: 0,
    columnSpan: 1,
    totalColumns: 4,
    rowFormatting: undefined,
    cellFormatting: undefined,
    ...overrides,
  })?.rPr;

describe("table conditional formatting coordinates", () => {
  test("explicit corner flags override coordinate-derived corners", () => {
    expect(resolve()?.bold).toBe(true);
    expect(resolve({ cellFormatting: { conditionalFormat: { nwCell: false } } })?.bold).toBeUndefined();
    expect(
      resolve({
        rowIndex: 1,
        columnIndex: 2,
        cellFormatting: { conditionalFormat: { nwCell: true } },
      })?.bold,
    ).toBe(true);
  });

  test("grid omissions and colspan keep band and last-column coordinates distinct", () => {
    const formatting = resolve({
      look: { lastColumn: true },
      rowIndex: 1,
      columnIndex: 1,
      columnSpan: 2,
      totalColumns: 4,
    });
    expect(formatting?.strike).toBeUndefined();
    expect(formatting?.smallCaps).toBe(true);

    expect(
      resolve({
        look: { lastColumn: true },
        rowIndex: 1,
        columnIndex: 1,
        columnSpan: 3,
        totalColumns: 4,
      })?.strike,
    ).toBe(true);
  });
});
