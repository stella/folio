import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "@stll/folio-core/server";
import {
  compareBatchOrder,
  orderBatchOperations,
  type BatchOrderAddress,
} from "@stll/docx-core/ops";

import { openReviewer, packDocument } from "../support/documents.ts";
import { sequentialGroups } from "../support/metamorphic.ts";
import type { Row } from "../support/oracle.ts";

const MODES = ["direct", "tracked-changes"] as const;
const ORDERS = [
  [0, 1, 2],
  [0, 2, 1],
  [1, 0, 2],
  [1, 2, 0],
  [2, 0, 1],
  [2, 1, 0],
] as const;

const view = (reviewer: Awaited<ReturnType<typeof openReviewer>>) =>
  reviewer.getContent().map(({ text, styleId, headingLevel }) => ({ text, styleId, headingLevel }));

type MixOptions = {
  mode: (typeof MODES)[number];
  order: (typeof ORDERS)[number];
  mirrored: boolean;
  inserted: string;
};

/** Apply a batch and the contract's right-to-left replay to independent copies. */
const compareMix = async ({ mode, order, mirrored, inserted }: MixOptions) => {
  const fixture = await packDocument(
    fromMarkdown(["Alpha.", "Bravo.", "Charlie.", "Delta.", "Echo."].join("\n\n")),
  );
  const batch = await openReviewer(fixture);
  const sequential = await openReviewer(fixture);
  const rows = batch.getContent() as Row[];
  const block = (index: number) => {
    const found = rows.at(index);
    assert.ok(found);
    return found;
  };
  const mix: FolioDocumentOperation[] = mirrored
    ? [
        { id: "insert", type: "insertAfterBlock", blockId: block(2).id, text: inserted },
        { id: "merge", type: "mergeBlockWithNext", blockId: block(0).id, separator: " " },
        { id: "replace", type: "replaceBlock", blockId: block(3).id, text: "Replaced delta." },
      ]
    : [
        { id: "insert", type: "insertAfterBlock", blockId: block(0).id, text: inserted },
        { id: "merge", type: "mergeBlockWithNext", blockId: block(2).id, separator: " " },
        { id: "replace", type: "replaceBlock", blockId: block(1).id, text: "Replaced bravo." },
      ];
  const operations = order.flatMap((index) => mix.at(index) ?? []);
  const request = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations,
  } as const;
  const result = batch.applyDocumentOperations(request, { undefinedReferences: "refuse" });
  assert.deepEqual(result.skipped, []);

  const groups = orderBatchOperations(
    operations.map((operation, inputIndex) => ({ operation, inputIndex })),
    ({ operation, inputIndex }) => {
      const anchor = rows.findIndex(({ id }) => id === operation.blockId);
      const position =
        operation.type === "mergeBlockWithNext" || operation.type === "insertAfterBlock"
          ? anchor + 1
          : anchor;
      return { type: "content" as const, position, rank: 1, inputIndex };
    },
  ).map(({ operation }) => [operation]);
  if (!mirrored && order.join(",") === "0,1,2") {
    assert.deepEqual(
      groups.flat().map(({ id }) => id),
      ["merge", "replace", "insert"],
      "the replay must not hoist the left insertion ahead of right-hand edits",
    );
  }
  for (const group of groups) {
    const replayOperations = group.flatMap(({ id }) =>
      operations.filter((operation) => operation.id === id),
    );
    const replayed = sequential.applyDocumentOperations(
      { ...request, operations: replayOperations },
      { undefinedReferences: "refuse" },
    );
    assert.deepEqual(replayed.skipped, []);
  }
  if (mode === "tracked-changes") {
    batch.acceptAll();
    sequential.acceptAll();
  }
  assert.deepEqual(view(sequential), view(batch));
};

test("pinned adjacent insertion, merge and replacement batches replay from pre-batch positions", async () => {
  // Fuzz seed 1101252753 (run 128 from base seed 1101252625).
  for (const mode of MODES) {
    await compareMix({ mode, order: ORDERS[0], mirrored: false, inserted: "Added 👍🏽 clause." });
  }
});

test("the batch comparator defines one total order", () => {
  const addresses: BatchOrderAddress[] = [];
  for (let inputIndex = 0; inputIndex < 4; inputIndex++) {
    for (let position = 0; position < 3; position++) {
      addresses.push({ type: "annotation", position, inputIndex });
      for (let rank = 0; rank < 3; rank++) {
        addresses.push({ type: "content", position, rank, inputIndex });
      }
      addresses.push({ type: "cell", tablePosition: position, bottom: 1, right: 2, inputIndex });
      addresses.push({
        type: "column",
        tablePosition: position,
        columnIndex: 1,
        insertion: inputIndex % 2 === 0,
        inputIndex,
      });
    }
  }
  const sign = (value: number) => Math.sign(value);
  for (const left of addresses) {
    for (const right of addresses) {
      const comparison = compareBatchOrder(left, right);
      const reverse = sign(compareBatchOrder(right, left));
      assert.ok(sign(comparison) === -reverse || (comparison === 0 && reverse === 0));
      if (comparison === 0) assert.deepEqual(left, right);
      for (const last of addresses) {
        if (comparison <= 0 && compareBatchOrder(right, last) <= 0) {
          assert.ok(compareBatchOrder(left, last) <= 0);
        }
      }
    }
  }
});

test("deletion precedes a table insertion through every three-operation permutation", async () => {
  const outcomes: ReturnType<typeof view>[] = [];
  for (const order of ORDERS) {
    const fixture = await packDocument(fromMarkdown("Alpha bravo."));
    const reviewer = await openReviewer(fixture);
    const original = reviewer.getContent() as Row[];
    const anchor = original.at(0);
    assert.ok(anchor);
    const split = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [{ id: "split", type: "splitBlock", blockId: anchor.id, offset: 5 }],
      },
      { undefinedReferences: "refuse" },
    );
    assert.deepEqual(split.skipped, []);
    const splitAnchor = (reviewer.getContent() as Row[]).at(0);
    assert.ok(splitAnchor);
    const operations = [
      {
        id: "table",
        type: "insertTable" as const,
        blockId: splitAnchor.id,
        rows: [["Term", "Value"]],
      },
      {
        id: "before",
        type: "insertBeforeBlock" as const,
        blockId: splitAnchor.id,
        text: "Before.",
      },
      { id: "delete", type: "deleteBlock" as const, blockId: splitAnchor.id },
    ];
    const orderedReplay = sequentialGroups({
      applied: order.flatMap((index) => operations.at(index) ?? []),
      preRows: reviewer.getContent() as Row[],
      mode: "tracked-changes",
    })
      .flat()
      .map(({ id }) => id);
    assert.ok(orderedReplay.indexOf("delete") < orderedReplay.indexOf("table"));
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: order.flatMap((index) => operations.at(index) ?? []),
      },
      { undefinedReferences: "refuse" },
    );
    assert.deepEqual(result.skipped, []);
    reviewer.acceptAll();
    outcomes.push(view(reviewer));
  }
  for (const outcome of outcomes.slice(1)) assert.deepEqual(outcome, outcomes.at(0));
});

test("generated adjacent insertion, merge and replacement mixes equal contract replay", async () => {
  const insertedTexts = ["Added clause.", "Added 👍🏽 clause.", "Cláusula añadida.", "סעיף נוסף."];
  for (const [index, order] of ORDERS.entries()) {
    await compareMix({
      mode: MODES[index % MODES.length] ?? "direct",
      order,
      mirrored: index % 2 === 1,
      inserted: insertedTexts[index % insertedTexts.length] ?? "Added clause.",
    });
  }
});
