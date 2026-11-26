import { describe, expect, test } from "bun:test";

import {
  buildTableDocx,
  readReviewerTables,
  tableReadingProblems,
  type TableSpec,
} from "../__tests__/tableOperationDocument";
import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";

// Two rows, the middle column merged down both.
const MERGED: TableSpec = {
  rows: 2,
  columns: 3,
  cells: [
    { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "n0" },
    { row: 0, column: 1, rowSpan: 2, columnSpan: 1, text: "n1" },
    { row: 0, column: 2, rowSpan: 1, columnSpan: 1, text: "n2" },
    { row: 1, column: 0, rowSpan: 1, columnSpan: 1, text: "n3" },
    { row: 1, column: 2, rowSpan: 1, columnSpan: 1, text: "n4" },
  ],
};

const NESTED: TableSpec = {
  rows: 1,
  columns: 1,
  cells: [
    { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "", nested: MERGED },
  ],
};

const blockIdOf = (reviewer: FolioDocxReviewer, text: string): string =>
  reviewer.getContent().find((block) => block.text === text)?.id ?? "";

const apply = async (
  spec: TableSpec,
  mode: "direct" | "tracked-changes",
  operations: (reviewer: FolioDocxReviewer) => FolioDocumentOperation[],
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await buildTableDocx(spec),
    { author: "Tester", session: "canonical" },
  );
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: operations(reviewer),
    },
    { undefinedReferences: "refuse" },
  );
  return { reviewer, result };
};

describe.each([
  ["top-level", MERGED],
  ["nested", NESTED],
] as const)("in a %s table", (_, spec) => {
  // The row inserted inside the merge continues it; the row deleted below
  // takes its own continuation only, so the merge goes on into the new row.
  test("a row inserted inside a merge whose last row the batch deletes continues it", async () => {
    const operations = (reviewer: FolioDocxReviewer) => [
      {
        id: "insert",
        type: "insertTableRow" as const,
        blockId: blockIdOf(reviewer, "n0"),
        position: "after" as const,
      },
      {
        id: "delete",
        type: "deleteTableRow" as const,
        blockId: blockIdOf(reviewer, "n3"),
      },
    ];
    const direct = await apply(spec, "direct", operations);
    expect(direct.result.skipped).toEqual([]);
    const directReading = await readReviewerTables(direct.reviewer);
    expect(tableReadingProblems(directReading)).toEqual([]);
    expect(
      directReading.snapshot.at(-1)?.cells.find(({ text }) => text === "n1")
        ?.rowSpan,
    ).toBe(2);

    const tracked = await apply(spec, "tracked-changes", operations);
    expect(tracked.result.skipped).toEqual([]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await tracked.reviewer.toBuffer(),
      { author: "Tester", session: "canonical" },
    );
    reopened.acceptAll();
    const accepted = await readReviewerTables(reopened);
    expect(tableReadingProblems(accepted)).toEqual([]);
    expect(accepted.snapshot).toEqual(directReading.snapshot);
  });
});
