import { describe, expect, test } from "bun:test";

import { compareContent } from "./content";
import type { FolioContentInputBlock } from "./content-types";

const block = (id: string, text: string): FolioContentInputBlock => ({
  identity: { type: "authoritative", id },
  kind: "paragraph",
  text,
});

const tableBlock = ({
  id,
  text,
  rowIndex,
  paragraphIndex,
  cellIndex = 0,
  containerId = "logical-cell",
}: {
  id: string;
  text: string;
  rowIndex: number;
  paragraphIndex: number;
  cellIndex?: number;
  containerId?: string;
}): FolioContentInputBlock => ({
  identity: { type: "authoritative", id },
  kind: "paragraph",
  text,
  containerPath: [
    { kind: "cell", identity: { type: "authoritative", id: containerId } },
  ],
  table: {
    outerTableIdentity: { type: "positional", id: "outer-table-0" },
    tableIdentity: { type: "positional", id: "table-0" },
    rowIdentity: { type: "positional", id: `row-${String(rowIndex)}` },
    cellIdentity: { type: "authoritative", id: containerId },
    outerTableIndex: 0,
    tableIndex: 0,
    rowIndex,
    cellIndex,
    gridColumnIndex: cellIndex,
    columnSpan: 1,
    rowSpan: 1,
    paragraphIndex,
  },
});

const successfulComparison = (
  base: FolioContentInputBlock[],
  revised: FolioContentInputBlock[],
) => {
  const result = compareContent({ base: { blocks: base }, revised: { blocks: revised } });
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
};

describe("neutral move scope", () => {
  test.each([
    {
      label: "exact text",
      baseText: "alpha beta gamma delta epsilon",
      revisedText: "alpha beta gamma delta epsilon",
    },
    {
      label: "edited text",
      baseText: "alpha beta gamma delta epsilon",
      revisedText: "alpha beta gamma delta zeta",
    },
  ])(
    "a same-slot stable-id replacement with $label is not a relocation",
    ({ baseText, revisedText }) => {
      const base = block("old-stable-id", baseText);
      const revised = block("new-stable-id", revisedText);

      const comparison = successfulComparison([base], [revised]);

      expect(comparison.events.map(({ type }) => type)).toEqual(["inserted", "deleted"]);
    },
  );

  test("a stable paragraph relocation survives raw row-coordinate shifts", () => {
    const baseMoved = tableBlock({
      id: "moved",
      text: "Title",
      rowIndex: 1,
      paragraphIndex: 0,
    });
    const revisedMoved = tableBlock({
      id: "moved",
      text: "Updated",
      rowIndex: 2,
      paragraphIndex: 2,
    });
    const base = [
      tableBlock({
        id: "header",
        text: "Header row",
        rowIndex: 0,
        paragraphIndex: 0,
        containerId: "header-cell",
      }),
      baseMoved,
      tableBlock({
        id: "anchor-a",
        text: "First durable anchor text",
        rowIndex: 1,
        paragraphIndex: 1,
      }),
      tableBlock({
        id: "anchor-b",
        text: "Second durable anchor text",
        rowIndex: 1,
        paragraphIndex: 2,
      }),
    ];
    const revised = [
      tableBlock({
        id: "inserted-row",
        text: "Inserted unrelated row",
        rowIndex: 0,
        paragraphIndex: 0,
        containerId: "inserted-cell",
      }),
      tableBlock({
        id: "header",
        text: "Header row",
        rowIndex: 1,
        paragraphIndex: 0,
        containerId: "header-cell",
      }),
      tableBlock({
        id: "anchor-a",
        text: "First durable anchor text",
        rowIndex: 2,
        paragraphIndex: 0,
      }),
      tableBlock({
        id: "anchor-b",
        text: "Second durable anchor text",
        rowIndex: 2,
        paragraphIndex: 1,
      }),
      revisedMoved,
    ];

    const comparison = successfulComparison(base, revised);
    const movedFrom = comparison.events.find(({ type }) => type === "movedFrom");
    const movedTo = comparison.events.find(({ type }) => type === "movedTo");

    expect(movedFrom).toMatchObject({
      type: "movedFrom",
      move: {
        relation: {
          base: {
            block: { identity: { id: "moved" }, table: { rowIndex: 1, cellIndex: 0 } },
          },
        },
      },
    });
    expect(movedTo).toMatchObject({
      type: "movedTo",
      move: {
        relation: {
          revised: {
            block: { identity: { id: "moved" }, table: { rowIndex: 2, cellIndex: 0 } },
          },
        },
      },
    });
    expect(movedTo?.type === "movedTo" ? movedTo.move : null).toBe(
      movedFrom?.type === "movedFrom" ? movedFrom.move : null,
    );
    const structural = comparison.events.filter(({ type }) => type === "structural");
    expect(structural).toHaveLength(1);
    expect(structural.at(0)).toMatchObject({
      type: "structural",
      memberIndex: 0,
      change: {
        type: "table-row-insert",
        tableIndex: 0,
        rowIndex: 0,
        blocks: [{ identity: { id: "inserted-row" } }],
      },
    });
  });
});
