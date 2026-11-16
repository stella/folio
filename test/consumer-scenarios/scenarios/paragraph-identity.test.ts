import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";

import { openReviewer, packDocument } from "../support/documents.ts";
import { createParagraphIdentityRemapping } from "../support/paragraph-identity.ts";

test("public join receipts remap chains while restored and deliberately stale ids keep their meaning", async () => {
  const reviewer = await openReviewer(
    await packDocument(fromMarkdown("Alpha.\n\nBravo.\n\nCharlie.\n\nTail.")),
  );
  const original = reviewer.getContent().at(0)?.id;
  assert.ok(original);
  const identities = createParagraphIdentityRemapping();
  const liveIds = () => new Set(reviewer.getContent().map(({ id }) => id));
  const merge = (id: string, blockId: string) =>
    reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        operations: [
          { id, type: "mergeBlockWithNext", blockId, separator: " " },
        ],
      },
      { undefinedReferences: "refuse" },
    );
  const first = merge("first", original);
  assert.deepEqual(first.skipped, []);
  const identity = first.applied.at(0)?.paragraphIdentity;
  assert.ok(identity?.type === "joined");
  assert.deepEqual(first.receipts.at(0)?.affected, [
    {
      type: "block",
      story: "main",
      blockId: identity.blockId,
      effect: "updated",
    },
    {
      type: "block",
      story: "main",
      blockId: identity.retiredBlockId,
      effect: "deleted",
    },
  ]);
  identities.record({ applied: first.applied, mode: "direct" });
  const middle = identities.resolve(original, liveIds());
  assert.notEqual(middle, original);
  const second = merge("second", middle);
  assert.deepEqual(second.skipped, []);
  identities.record({ applied: second.applied, mode: "direct" });
  const survivor = identities.resolve(original, liveIds());
  assert.notEqual(survivor, middle);
  assert.equal(identities.resolve(middle, liveIds()), survivor);
  assert.equal(
    reviewer.getContent().find(({ id }) => id === survivor)?.text,
    "Alpha. Bravo. Charlie.",
  );

  // Remapping is an explicit caller action. A negative test that sends the
  // retired id as written must still exercise the public missing-id refusal.
  assert.deepEqual(
    merge("stale", original).skipped.map(({ reason }) => reason),
    ["missingBlock"],
  );
  assert.ok(second.undoHandle);
  assert.equal(
    reviewer.undoDocumentOperations(second.undoHandle).status,
    "undone",
  );
  assert.equal(identities.resolve(original, liveIds()), middle);
  assert.equal(identities.resolve(middle, liveIds()), middle);
});
