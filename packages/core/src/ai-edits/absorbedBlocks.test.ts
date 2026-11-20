/**
 * A canonical reviewer names, for every paragraph a resolution or a batch
 * joins into another, the paragraph that holds its content now: a join keeps
 * the second paragraph's id.
 */
import { expect, test } from "bun:test";

import { textRun } from "../__tests__/operationBatchDocuments";
import { createDocx } from "../docx/rezip";
import { FolioDocxReviewer } from "./headless";

test("accepting all names each joined paragraph's survivor", async () => {
  const docx = await createDocx({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "10000001",
            content: [textRun("One ")],
            pPrMark: { kind: "del", info: { id: 1, author: "A" } },
          },
          {
            type: "paragraph",
            paraId: "10000002",
            content: [textRun("Two ")],
            pPrMark: { kind: "del", info: { id: 2, author: "A" } },
          },
          {
            type: "paragraph",
            paraId: "10000003",
            content: [textRun("Three")],
          },
        ],
      },
    },
  });
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
    session: "canonical",
  });
  const resolution = reviewer.acceptAll();
  expect(resolution.resolved).toBeGreaterThan(0);
  expect(
    resolution.absorbedBlocks?.toSorted((left, right) =>
      left.blockId.localeCompare(right.blockId),
    ),
  ).toEqual([
    { blockId: "10000001", survivorBlockId: "10000003" },
    { blockId: "10000002", survivorBlockId: "10000003" },
  ]);
  expect(reviewer.getContent().map(({ id }) => id)).toEqual(["10000003"]);
});

test("rejecting all reports nothing where no paragraph joins", async () => {
  const docx = await createDocx({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "10000001",
            content: [textRun("One")],
            pPrMark: { kind: "del", info: { id: 1, author: "A" } },
          },
          { type: "paragraph", paraId: "10000002", content: [textRun("Two")] },
        ],
      },
    },
  });
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
    session: "canonical",
  });
  expect(reviewer.rejectAll().absorbedBlocks).toBeUndefined();
});

test("a tracked deletion of one's own inserted paragraph names the paragraph it joined", async () => {
  const docx = await createDocx({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "10000000",
            content: [textRun("First")],
          },
          {
            type: "paragraph",
            paraId: "10000001",
            content: [textRun("Added")],
            pPrMark: { kind: "ins", info: { id: 1, author: "Reviewer" } },
          },
          { type: "paragraph", paraId: "10000002", content: [textRun("Kept")] },
        ],
      },
    },
  });
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
    author: "Reviewer",
    session: "canonical",
  });
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: "10000001" }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied.at(0)?.paragraphIdentity).toEqual({
    type: "joined",
    blockId: "10000002",
    retiredBlockId: "10000001",
  });
});
