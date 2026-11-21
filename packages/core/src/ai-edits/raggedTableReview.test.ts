import { expect, test } from "bun:test";

import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { TableCell } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

let next = 0x20000000;
const cell = (
  text: string,
  change?: "tableCellInsertion" | "tableCellDeletion",
): TableCell => ({
  type: "tableCell",
  ...(change === undefined
    ? {}
    : {
        structuralChange: {
          type: change,
          info: {
            id: change === "tableCellInsertion" ? 11 : 12,
            author: "Reviewer",
            date: "2026-11-20T00:00:00Z",
          },
        },
      }),
  content: [
    {
      type: "paragraph",
      paraId: (next++).toString(16).toUpperCase(),
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

/** A one-column grid whose middle row holds an inserted and a deleted cell. */
const raggedDocx = () => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "table",
      columnWidths: [4000],
      rows: [
        { type: "tableRow", cells: [cell("A")] },
        {
          type: "tableRow",
          cells: [
            cell("B", "tableCellInsertion"),
            cell("C", "tableCellDeletion"),
          ],
        },
        { type: "tableRow", cells: [cell("D")] },
      ],
    },
    { type: "paragraph", paraId: "2FFFFFFF", content: [] },
  ];
  return createDocx(document);
};

for (const decision of ["accept", "reject"] as const)
  test(`a ragged table's cell changes ${decision} on the canonical session`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await raggedDocx(), {
      session: "canonical",
    });
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    expect(reviewer.getChanges()).toEqual([]);
    const saved = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const table = saved.package.document.content.find(
      (block) => block.type === "table",
    );
    if (table?.type !== "table") throw new Error("expected the table");
    // The cell that stays takes the removed cell's column.
    const middle = table.rows.at(1)?.cells ?? [];
    expect(middle).toHaveLength(1);
    expect(middle.at(0)?.formatting?.gridSpan).toBe(2);
    expect(JSON.stringify(middle.at(0)?.content)).toContain(
      decision === "accept" ? '"B"' : '"C"',
    );
  });
