import { expect, test } from "bun:test";
import assert from "node:assert/strict";
import { buildDocx } from "../../packages/cli/src/__tests__/fixtures";
import {
  FolioDocxReviewer,
  type FolioDocumentOperationBatch,
} from "../../packages/core/src/server";
import { assertGeneratedFlowReceipt } from "./flow-receipt";

const cases = [
  { blockId: "10000001", heading: true, text: "café" },
  { blockId: "10000002", heading: false, text: "東京" },
  { blockId: "10000001", heading: true, text: "e\u0301" },
] as const;

test.each(cases)(
  "accepts correct insert and replace receipts at $blockId for $text",
  async ({ blockId, heading, text }) => {
    const source = await buildDocx([
      { text: "token0 original clause", paraId: "10000001", ...(heading && { style: "Heading1" }) },
      { text: "token1 original clause", paraId: "10000002" },
    ]);
    const reviewer = await FolioDocxReviewer.fromBuffer(new Uint8Array(source).buffer);
    const operations = [
      { id: "insert", type: "insertAfterBlock", blockId, text },
      {
        id: "replace",
        type: "replaceInBlock",
        blockId,
        find: blockId === "10000001" ? "token0" : "token1",
        replace: `${blockId === "10000001" ? "token0" : "token1"} ${text}`,
      },
    ] as const satisfies readonly FolioDocumentOperationBatch["operations"][number][];

    for (const operation of operations) {
      const batch = {
        version: 1,
        mode: "direct",
        operations: [operation],
      } as const satisfies FolioDocumentOperationBatch;
      assertGeneratedFlowReceipt(reviewer.applyDocumentOperations(batch), operation);
    }
  },
);

test("rejects insertion receipts with the wrong discriminator or anchor", async () => {
  const source = await buildDocx([
    { text: "token0 original clause", paraId: "10000001", style: "Heading1" },
    { text: "token1 original clause", paraId: "10000002" },
  ]);
  const reviewer = await FolioDocxReviewer.fromBuffer(new Uint8Array(source).buffer);
  const operation = {
    id: "insert",
    type: "insertAfterBlock",
    blockId: "10000001",
    text: "café",
  } as const satisfies FolioDocumentOperationBatch["operations"][number];
  const batch = {
    version: 1,
    mode: "direct",
    operations: [operation],
  } as const satisfies FolioDocumentOperationBatch;
  const result = reviewer.applyDocumentOperations(batch);
  const wrongType = structuredClone(result);
  const wrongTypeReceipt = wrongType.receipts.at(0);
  assert.ok(wrongTypeReceipt);
  wrongTypeReceipt.affected[0] = {
    type: "block",
    story: "main",
    blockId: operation.blockId,
    effect: "updated",
  };
  const wrongAnchor = structuredClone(result);
  const wrongAnchorReceipt = wrongAnchor.receipts.at(0);
  assert.ok(wrongAnchorReceipt);
  wrongAnchorReceipt.affected[0] = {
    type: "insertion",
    story: "main",
    anchorBlockId: "10000002",
    position: "after",
    content: "block",
  };

  expect(() => assertGeneratedFlowReceipt(wrongType, operation)).toThrow(
    "Insert operation receipt does not describe its insertion",
  );
  expect(() => assertGeneratedFlowReceipt(wrongAnchor, operation)).toThrow(
    "Insert operation receipt does not describe its insertion",
  );
});
