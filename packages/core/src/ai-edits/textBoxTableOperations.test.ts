/**
 * Table operations on a table inside a text box, addressed by the ids the
 * reviewer shows for the box's paragraphs, on both sessions.
 */
import { describe, expect, test } from "bun:test";

import {
  buildTextBoxTableDocument,
  findTextBoxShape,
} from "../__tests__/textBoxTableDocument";
import type { FolioDocumentOperation } from "../document-operations";
import { parseDocx } from "../docx/parser";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer
    .snapshot()
    .blocks.find((candidate) => candidate.text.includes(text));
  if (block === undefined) throw new Error(`no block holds ${text}`);
  return block.id;
};

const boxTable = async (reviewer: FolioDocxReviewer) => {
  const table = findTextBoxShape(
    await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    }),
  ).textBody?.content.find(({ type }) => type === "table");
  if (table?.type !== "table") throw new Error("the box lost its table");
  return table;
};

type ApplyOptions = {
  reviewer: FolioDocxReviewer;
  mode: "direct" | "tracked-changes";
  operation: FolioDocumentOperation;
};

const apply = ({ reviewer, mode, operation }: ApplyOptions) =>
  reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [operation] },
    { undefinedReferences: "refuse" },
  );

describe.each(SESSIONS)("a table in a text box (%s)", (session) => {
  const open = async () =>
    FolioDocxReviewer.fromBuffer(await buildTextBoxTableDocument(), {
      author: "Reviewer",
      session,
    });

  test("takes a tracked row and keeps it once accepted", async () => {
    const reviewer = await open();
    const result = apply({
      reviewer,
      mode: "tracked-changes",
      operation: {
        id: "row",
        type: "insertTableRow",
        blockId: blockId(reviewer, "Cell value"),
        cellTexts: ["Added value"],
      },
    });
    expect(result.skipped).toEqual([]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    expect(reopened.getChanges().map(({ type }) => type)).toContain(
      "rowInserted",
    );
    reopened.acceptAll();
    expect(reopened.getChanges()).toEqual([]);
    expect((await boxTable(reopened)).rows).toHaveLength(2);
  });

  test("takes a tracked column and drops it once rejected", async () => {
    const reviewer = await open();
    const result = apply({
      reviewer,
      mode: "tracked-changes",
      operation: {
        id: "column",
        type: "insertTableColumn",
        blockId: blockId(reviewer, "Cell value"),
        cellTexts: ["Added value"],
      },
    });
    expect(result.skipped).toEqual([]);
    expect((await boxTable(reviewer)).rows.at(0)?.cells).toHaveLength(2);
    reviewer.rejectAll();
    expect(reviewer.getChanges()).toEqual([]);
    expect((await boxTable(reviewer)).rows.at(0)?.cells).toHaveLength(1);
  });

  test("merges the cells of a column it inserted", async () => {
    const reviewer = await open();
    expect(
      apply({
        reviewer,
        mode: "direct",
        operation: {
          id: "column",
          type: "insertTableColumn",
          blockId: blockId(reviewer, "Cell value"),
          cellTexts: ["Added value"],
        },
      }).skipped,
    ).toEqual([]);
    expect(
      apply({
        reviewer,
        mode: "direct",
        operation: {
          id: "merge",
          type: "mergeTableCells",
          blockId: blockId(reviewer, "Cell value"),
          endBlockId: blockId(reviewer, "Added value"),
        },
      }).skipped,
    ).toEqual([]);
    expect((await boxTable(reviewer)).rows.at(0)?.cells).toHaveLength(1);
  });
});
