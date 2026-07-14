import { describe, expect, test } from "bun:test";

import { schema } from "../../prosemirror/schema";
import { toFlowBlocks } from "./toFlowBlocks";

describe("table cell diagonal border conversion", () => {
  test("carries visible diagonals into the layout model", () => {
    const cell = schema.node(
      "tableCell",
      {
        borders: {
          topLeftToBottomRight: {
            color: { rgb: "123456" },
            size: 12,
            style: "dashed",
          },
          topRightToBottomLeft: { style: "nil" },
        },
      },
      [schema.node("paragraph", null, [schema.text("Cell")])],
    );
    const doc = schema.node("doc", null, [
      schema.node("table", { columnWidths: [1200] }, [schema.node("tableRow", null, [cell])]),
    ]);

    const table = toFlowBlocks(doc).find((block) => block.kind === "table");
    expect(table?.kind).toBe("table");
    if (!table || table.kind !== "table") {
      return;
    }

    expect(table.rows.at(0)?.cells.at(0)?.borders?.topLeftToBottomRight).toEqual({
      color: "#123456",
      style: "dashed",
      width: 2,
    });
    expect(table.rows.at(0)?.cells.at(0)?.borders?.topRightToBottomLeft).toBeUndefined();
  });
});
