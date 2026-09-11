import { describe, expect, test } from "bun:test";

import { alignTableColumns } from "./column-alignment";
import { contentBlockFixture, tableLocationFixture } from "./content-test-fixtures";
import type { FolioContentBlock } from "./content-types";

type TableBlockOptions = {
  id: string;
  text: string;
  rowIndex: number;
  cellIndex: number;
  gridColumnIndex: number;
  columnSpan?: number;
};

const countedTableBlock = ({
  id,
  text,
  rowIndex,
  cellIndex,
  gridColumnIndex,
  columnSpan = 1,
}: TableBlockOptions): { block: FolioContentBlock; reads: () => number } => {
  let textReads = 0;
  const captured = contentBlockFixture(id, text, {
    identityType: "persistent-hint",
    table: tableLocationFixture({
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex,
      cellIndex,
      gridColumnIndex,
      columnSpan,
      rowSpan: 1,
      paragraphIndex: 0,
    }),
  });
  const block = {
    ...captured,
    get text() {
      textReads++;
      return text;
    },
  } satisfies FolioContentBlock;
  return { block, reads: () => textReads };
};

describe("bounded table-column alignment", () => {
  test("a long cell spanning all 63 columns contributes its text once", () => {
    const longText = "long table cell ".repeat(16_384);
    const base = countedTableBlock({
      id: "base-wide",
      text: longText,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 62,
    });
    const revised = countedTableBlock({
      id: "revised-wide",
      text: longText,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 63,
    });

    expect(alignTableColumns([base.block], [revised.block])).toBeNull();
    expect(base.reads()).toBe(1);
    expect(revised.reads()).toBe(1);
  });

  test("an equal-width million-row sparse grid falls back before reading signatures", () => {
    const base = countedTableBlock({
      id: "base-sparse",
      text: "Base",
      rowIndex: 999_999,
      cellIndex: 0,
      gridColumnIndex: 0,
    });
    const revised = countedTableBlock({
      id: "revised-sparse",
      text: "Revised",
      rowIndex: 999_999,
      cellIndex: 0,
      gridColumnIndex: 0,
    });

    expect(alignTableColumns([base.block], [revised.block])).toBeNull();
    expect(base.reads()).toBe(0);
    expect(revised.reads()).toBe(0);
  });

  test("an unsafe derived row boundary is rejected", () => {
    const rowIndex = Number.MAX_SAFE_INTEGER;
    const rowSpan = 1;
    const block = contentBlockFixture("unsafe-row-end", "Unsafe row boundary", {
      identityType: "persistent-hint",
      table: tableLocationFixture({
        outerTableIndex: 0,
        tableIndex: 0,
        rowIndex,
        cellIndex: 0,
        gridColumnIndex: 0,
        columnSpan: 1,
        rowSpan,
        paragraphIndex: 0,
      }),
    });

    expect(Number.isSafeInteger(rowIndex)).toBe(true);
    expect(Number.isSafeInteger(rowSpan)).toBe(true);
    expect(Number.isSafeInteger(rowIndex + rowSpan)).toBe(false);
    expect(alignTableColumns([block], [block])).toBeNull();
  });
});
