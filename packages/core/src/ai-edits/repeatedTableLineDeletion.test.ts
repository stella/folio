/**
 * A batch that deletes one table row or column twice, through any two of its
 * cells, deletes it once: the repeat is a no-op, never a deletion of the row
 * or column that moved into its place. Every pair of cells of a row (or a
 * column) of a 3x3 table, on both engines, applied directly and tracked.
 */

import { describe, expect, test } from "bun:test";

import {
  buildTableDocx,
  readReviewerTables,
  type TableSpec,
} from "../__tests__/tableOperationDocument";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

const SIZE = 3;

const GRID: TableSpec = {
  rows: SIZE,
  columns: SIZE,
  cells: Array.from({ length: SIZE * SIZE }, (_, slot) => ({
    row: Math.floor(slot / SIZE),
    column: slot % SIZE,
    rowSpan: 1,
    columnSpan: 1,
    text: `r${Math.floor(slot / SIZE)}c${slot % SIZE}`,
  })),
};

const SESSIONS = ["canonical"] as const;
const MODES = ["direct", "tracked-changes"] as const;
const AXES = [
  {
    type: "deleteTableRow",
    cell: (line: number, at: number) => `r${line}c${at}`,
  },
  {
    type: "deleteTableColumn",
    cell: (line: number, at: number) => `r${at}c${line}`,
  },
] as const;

describe.each(SESSIONS)("%s session", (session) => {
  const open = (bytes: ArrayBuffer) =>
    FolioDocxReviewer.fromBuffer(bytes, { author: "Tester", session });

  describe.each(MODES)("%s", (mode) => {
    test.each(AXES)("$type twice deletes it once", async ({ type, cell }) => {
      const base = await buildTableDocx(GRID);
      const run = async (texts: readonly string[]) => {
        const reviewer = await open(base);
        const blockId = (text: string) =>
          reviewer.getContent().find((block) => block.text === text)?.id ??
          `missing ${text}`;
        const operations: FolioDocumentOperation[] = texts.map(
          (text, index) => ({
            id: `op-${index}`,
            type,
            blockId: blockId(text),
          }),
        );
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations,
          },
          { undefinedReferences: "refuse" },
        );
        const { texts: read, snapshot } = await readReviewerTables(reviewer);
        return { result, read, snapshot };
      };
      for (let line = 0; line < SIZE; line++)
        for (let first = 0; first < SIZE; first++)
          for (let second = 0; second < SIZE; second++) {
            const once = await run([cell(line, first)]);
            const twice = await run([cell(line, first), cell(line, second)]);
            const label = `${type} ${line} via ${first}, ${second}`;
            expect(
              once.result.applied.map(({ id }) => id),
              label,
            ).toEqual(["op-0"]);
            expect(
              twice.result.applied.map(({ id }) => id),
              label,
            ).toEqual(["op-0"]);
            expect(twice.result.skipped, label).toEqual([
              { id: "op-1", reason: "noopOperation" },
            ]);
            expect(twice.read, label).toEqual(once.read);
            expect(twice.snapshot, label).toEqual(once.snapshot);
          }
    });
  });
});
