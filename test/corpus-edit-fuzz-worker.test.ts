import assert from "node:assert/strict";
import { test } from "node:test";

import { createFolioAITextRangeHandle } from "@stll/folio-core/server";

import { openReviewer, plainDocument } from "./consumer-scenarios/support/documents.ts";
import { capture } from "./consumer-scenarios/support/oracle.ts";
import { coreBatch } from "./consumer-scenarios/support/operations.ts";
import { assertSavedOutcome } from "./corpus-edit-fuzz-worker.ts";

test("saved-only comment loss fails the edit outcome check", async () => {
  const original = await plainDocument();
  const reviewer = await openReviewer(original);
  const block = reviewer.getContent().find(({ text }) => text.includes("Supplier"));
  assert.ok(block);
  const startOffset = block.text.indexOf("Supplier");
  const operation = {
    type: "commentOnRange",
    range: createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset,
      endOffset: startOffset + "Supplier".length,
    }),
    comment: { text: "Review the supplier clause." },
  };
  const pre = await capture(reviewer, "direct");
  const receipt = reviewer.applyDocumentOperations(coreBatch([operation], "direct") as never);
  assert.equal(receipt.applied.length, 1);

  const saved = new Uint8Array(await reviewer.toBuffer());
  assert.deepEqual(
    await assertSavedOutcome({
      reopened: await openReviewer(saved),
      pre,
      operation,
      applied: true,
      context: "intact save",
    }),
    [],
  );

  // A faulty serializer could drop comments while preserving block text and kind.
  const corrupted = await openReviewer(original);
  assert.deepEqual(
    corrupted.getContent().map(({ text, kind }) => ({ text, kind })),
    reviewer.getContent().map(({ text, kind }) => ({ text, kind })),
  );
  await assert.rejects(
    assertSavedOutcome({
      reopened: corrupted,
      pre,
      operation,
      applied: true,
      context: "corrupted save",
    }),
    /comment/u,
  );
});

test("a refused edit is checked after save and reopen", async () => {
  const reviewer = await openReviewer(await plainDocument());
  const pre = await capture(reviewer, "direct");
  const operation = { type: "deleteBlock", blockId: "missing-block" };
  const receipt = reviewer.applyDocumentOperations(coreBatch([operation], "direct") as never);
  assert.equal(receipt.applied.length, 0);

  assert.deepEqual(
    await assertSavedOutcome({
      reopened: await openReviewer(new Uint8Array(await reviewer.toBuffer())),
      pre,
      operation,
      applied: false,
      context: "refused edit",
    }),
    [],
  );
});

test("saved-only revision loss fails the reject check", async () => {
  const original = await plainDocument();
  const tracked = await openReviewer(original);
  const block = tracked.getContent().find(({ text }) => text.includes("Supplier"));
  assert.ok(block);
  const operation = {
    type: "replaceInBlock",
    blockId: block.id,
    find: "Supplier",
    replace: "Provider",
  };
  const pre = await capture(tracked, "tracked-changes");
  const trackedReceipt = tracked.applyDocumentOperations(
    coreBatch([operation], "tracked-changes") as never,
  );
  assert.equal(trackedReceipt.applied.length, 1);
  assert.deepEqual(
    await assertSavedOutcome({
      reopened: await openReviewer(new Uint8Array(await tracked.toBuffer())),
      pre,
      operation,
      applied: true,
      context: "intact revisions",
    }),
    [],
  );

  // The accepted text survives, but a faulty serializer has flattened the revision.
  const flattened = await openReviewer(original);
  const directReceipt = flattened.applyDocumentOperations(
    coreBatch([operation], "direct") as never,
  );
  assert.equal(directReceipt.applied.length, 1);
  await assert.rejects(
    assertSavedOutcome({
      reopened: await openReviewer(new Uint8Array(await flattened.toBuffer())),
      pre,
      operation,
      applied: true,
      context: "flattened revisions",
    }),
    /rejecting every change/u,
  );
});
