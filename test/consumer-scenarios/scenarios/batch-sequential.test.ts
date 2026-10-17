import assert from "node:assert/strict";
import { test } from "node:test";

import fc from "fast-check";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "@stll/folio-core/server";

import { assertProperty } from "../../property-testing.ts";
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

  const groups = sequentialGroups({ applied: operations, preRows: rows, mode });
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

test("generated adjacent insertion, merge and replacement mixes equal contract replay", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.constantFrom(...MODES),
      fc.constantFrom(...ORDERS),
      fc.boolean(),
      fc.constantFrom("Added clause.", "Added 👍🏽 clause.", "Cláusula añadida.", "סעיף נוסף."),
      async (mode, order, mirrored, inserted) => {
        await compareMix({ mode, order, mirrored, inserted });
      },
    ),
    { numRuns: 24, seed: 1101252753, id: "adjacent insert merge replace contract replay" },
  );
});
