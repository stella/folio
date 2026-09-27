/** Synthetic reductions of model tool-call shapes, exercised through packed public exports. */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";
import { getFolioDocumentOperationIssues } from "@stll/folio-core/server";

import { openReviewer, plainDocument } from "../support/documents.ts";
import { assertHealthy, visibleState } from "../support/invariants.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Row = { blockId: string; kind: string; text: string; blockTextHash: string };
type Summary = {
  applied: { id: string }[];
  skipped: { id: string; reason: string }[];
  issues: { code: string }[];
  normalizations: { path: string; message: string }[];
};

const call = (reviewer: Reviewer, name: string, args: unknown) =>
  executeFolioToolCallUntyped(name, args, createReviewerBridge(reviewer, { mode: "direct" }), {});

const rowsOf = (reviewer: Reviewer): Row[] => {
  const result = call(reviewer, "read_document", {});
  assert.ok(result.ok, "read_document refused");
  return result.result as Row[];
};

const summaryOf = (reviewer: Reviewer, operations: unknown[]): Summary => {
  const result = call(reviewer, "suggest_changes", { operations });
  assert.ok(result.ok, result.ok ? "" : result.error);
  return result.result as Summary;
};

describe("observed model tool-call mistakes", () => {
  for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const) {
    test(`${type}: text plus hardPageBreak is refused before a later valid edit`, async () => {
      const reviewer = await openReviewer(await plainDocument());
      const before = visibleState(reviewer);
      const target = rowsOf(reviewer).find((row) => row.text.includes("Supplier delivers"));
      assert.ok(target);
      const result = call(reviewer, "suggest_changes", {
        operations: [
          {
            type,
            blockId: target.blockId,
            text: "An added sentence.",
            hardPageBreak: { clear: "none" },
            // A projected tool schema can also supply fields for other operation kinds.
            find: "Supplier",
            replace: "Provider",
          },
          {
            type: "replaceInBlock",
            blockId: target.blockId,
            find: "Supplier",
            replace: "Provider",
          },
        ],
      });
      assert.equal(result.ok, false, "a malformed first operation must refuse the whole call");
      if (result.ok) return;
      assert.match(result.error, /operations\[0\]\.text/u);
      assert.match(result.error, /omit hardPageBreak to insert text/u);
      assert.match(result.error, /set text to "" to insert a hard page break/u);
      assert.deepEqual(visibleState(reviewer), before, "the later valid operation must not land");
      await assertHealthy(reviewer, `${type} with text and hardPageBreak`);
    });
  }

  test("a stale block hash is refused, then a fresh read makes the retry apply", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const original = rowsOf(reviewer).find((row) => row.text.includes("Supplier delivers"));
    assert.ok(original);
    const first = summaryOf(reviewer, [
      {
        type: "replaceInBlock",
        blockId: original.blockId,
        find: "Supplier",
        replace: "Provider",
        precondition: { blockTextHash: original.blockTextHash },
      },
    ]);
    assert.equal(first.applied.length, 1);
    const beforeRetry = visibleState(reviewer);
    const stale = summaryOf(reviewer, [
      {
        type: "replaceInBlock",
        blockId: original.blockId,
        find: "Provider",
        replace: "Vendor",
        precondition: { blockTextHash: original.blockTextHash },
      },
    ]);
    assert.equal(stale.applied.length, 0);
    assert.deepEqual(
      stale.issues.map(({ code }) => code),
      ["preconditionFailed"],
    );
    assert.match(stale.skipped[0]?.reason ?? "", /re-read the document and retry/u);
    assert.deepEqual(visibleState(reviewer), beforeRetry);
    await assertHealthy(reviewer, "stale precondition refusal");

    const fresh = rowsOf(reviewer).find((row) => row.blockId === original.blockId);
    assert.ok(fresh);
    const repaired = summaryOf(reviewer, [
      {
        type: "replaceInBlock",
        blockId: fresh.blockId,
        find: "Provider",
        replace: "Vendor",
        precondition: { blockTextHash: fresh.blockTextHash },
      },
    ]);
    assert.equal(repaired.applied.length, 1);
    assert.equal(repaired.skipped.length, 0);
    assert.ok(rowsOf(reviewer).some((row) => row.text.includes("Vendor delivers")));
    await assertHealthy(reviewer, "fresh precondition retry");
  });

  test("fields for another operation kind are reported while the intended edit applies", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const target = rowsOf(reviewer).find((row) => row.text.includes("Supplier delivers"));
    assert.ok(target);
    const outcome = summaryOf(reviewer, [
      {
        type: "replaceInBlock",
        blockId: target.blockId,
        find: "Supplier",
        replace: "Provider",
        text: "This field belongs to another operation.",
        hardPageBreak: { clear: "none" },
        cellTexts: ["Unrelated cell"],
      },
    ]);
    assert.equal(outcome.applied.length, 1);
    assert.equal(outcome.skipped.length, 0);
    assert.deepEqual(
      outcome.normalizations.map(({ path }) => path),
      ["operations[0].text", "operations[0].hardPageBreak", "operations[0].cellTexts"],
    );
    assert.ok(rowsOf(reviewer).some((row) => row.text.includes("Provider delivers")));
    await assertHealthy(reviewer, "irrelevant projected fields");
  });

  test("an unsupported block refusal gives repair and stop guidance", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const target = rowsOf(reviewer).find((row) => row.text.includes("Supplier delivers"));
    assert.ok(target);
    const before = visibleState(reviewer);
    const bridge = createReviewerBridge(reviewer, { mode: "direct" });
    // A host may know a target has structure the current operation cannot change.
    const refusingBridge = {
      ...bridge,
      applyDocumentOperations: (batch: Parameters<typeof bridge.applyDocumentOperations>[0]) => {
        const skipped = batch.operations.map(({ id }) => ({
          id,
          reason: "unsupportedBlock" as const,
        }));
        return {
          version: batch.version,
          status: "rejected" as const,
          applied: [],
          skipped,
          issues: getFolioDocumentOperationIssues(batch.operations, skipped),
          receipts: [],
          undoHandle: null,
        };
      },
    };
    const response = executeFolioToolCallUntyped(
      "suggest_changes",
      {
        operations: [
          {
            type: "deleteBlock",
            blockId: target.blockId,
            precondition: { blockTextHash: target.blockTextHash },
          },
        ],
      },
      refusingBridge,
      {},
    );
    assert.ok(response.ok, response.ok ? "" : response.error);
    const result = response.result as Summary;
    assert.equal(result.applied.length, 0);
    assert.deepEqual(
      result.issues.map(({ code }) => code),
      ["unsupportedBlock"],
    );
    assert.match(result.skipped[0]?.reason ?? "", /read_document to inspect/u);
    assert.match(result.skipped[0]?.reason ?? "", /tell the user this edit cannot be applied/u);
    assert.deepEqual(visibleState(reviewer), before);
    await assertHealthy(reviewer, "unsupported block refusal");
  });
});
