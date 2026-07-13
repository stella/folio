import { describe, expect, test } from "bun:test";

import { getColumns } from "./sectionColumns";

describe("getColumns", () => {
  test("preserves authored unequal widths and per-column spacing", () => {
    const columns = getColumns({
      columnCount: 2,
      columnSpace: 300,
      equalWidth: false,
      columns: [{ width: 1500, space: 750 }, { width: 3000 }],
    });

    expect(columns).toEqual({
      count: 2,
      gap: 20,
      equalWidth: false,
      widths: [100, 200],
      gaps: [50],
    });
  });
});
