import { expect, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import { FolioDocxReviewer } from "./headless";

test.each(["legacy", "canonical"] as const)(
  "rejecting a terminal table insertion restores the original final paragraph (%s)",
  async (session) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await paragraphsDocx([[textRun("Final paragraph")]]),
      { session },
    );
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "append",
            type: "insertTable",
            blockId: "10000000",
            rows: [["Cell"]],
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    const saved = await reviewer.toBuffer();
    const rejected = await FolioDocxReviewer.fromBuffer(saved, { session });
    rejected.rejectAll();
    expect(rejected.toDocument().package.document.content).toMatchObject([
      {
        type: "paragraph",
        content: [textRun("Final paragraph")],
      },
    ]);
  },
);

test.each(["legacy", "canonical"] as const)(
  "rejecting a table appended beside a pending final deletion restores the source (%s)",
  async (session) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await paragraphsDocx([[textRun("Start")], [textRun("Final paragraph")]]),
      { session },
    );
    const options = { undefinedReferences: "refuse" } as const;
    const removed = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          { id: "delete", type: "deleteBlock", blockId: "10000001" },
        ],
      },
      options,
    );
    expect(removed.skipped).toEqual([]);
    const appended = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "append",
            type: "insertTable",
            blockId: "10000001",
            rows: [["Cell"]],
          },
        ],
      },
      options,
    );
    expect(appended.skipped).toEqual([]);
    const rejected = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    rejected.rejectAll();
    expect(rejected.toDocument().package.document.content).toMatchObject([
      { type: "paragraph", content: [textRun("Start")] },
      { type: "paragraph", content: [textRun("Final paragraph")] },
    ]);
  },
);
