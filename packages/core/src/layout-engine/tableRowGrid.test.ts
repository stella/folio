import { describe, expect, test } from "bun:test";

import { getTableRowLeadingWidth, type TableRow } from "./types";

const paragraph = {
  kind: "paragraph" as const,
  id: "paragraph",
  runs: [],
};

describe("table row leading grid width", () => {
  test("reserves omitted columns before ordinary cell content", () => {
    const row: TableRow = {
      id: "row",
      gridBefore: 1,
      cells: [{ id: "cell", blocks: [paragraph] }],
    };

    expect(getTableRowLeadingWidth(row, [40, 60])).toBe(40);
  });

  test("lets a nested table own its grid origin", () => {
    const row: TableRow = {
      id: "row",
      gridBefore: 1,
      cells: [
        {
          id: "cell",
          blocks: [{ kind: "table", id: "nested", rows: [] }, paragraph],
        },
      ],
    };

    expect(getTableRowLeadingWidth(row, [40, 60])).toBe(0);
  });
});
