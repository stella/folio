import assert from "node:assert/strict";
import { test } from "node:test";

import { createFolioAITextRangeHandle } from "@stll/folio-core/server";

import {
  mergedTableDocument,
  openReviewer,
  plainDocument,
} from "./consumer-scenarios/support/documents.ts";
import { coreBatch } from "./consumer-scenarios/support/operations.ts";
import {
  capture,
  withLiveMintedIds,
  withoutMintedIds,
} from "./consumer-scenarios/support/oracle.ts";
import {
  assertSavedOutcome,
  compareSdk,
  runSteps,
} from "./corpus-edit-fuzz-worker.ts";

test("saved-only comment loss fails the edit outcome check", async () => {
  const original = await plainDocument();
  const reviewer = await openReviewer(original);
  const block = reviewer
    .getContent()
    .find(({ text }) => text.includes("Supplier"));
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
  const receipt = reviewer.applyDocumentOperations(
    coreBatch([operation], "direct") as never,
    {
      undefinedReferences: "refuse",
    },
  );
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
  const receipt = reviewer.applyDocumentOperations(
    coreBatch([operation], "direct") as never,
    {
      undefinedReferences: "refuse",
    },
  );
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
  const block = tracked
    .getContent()
    .find(({ text }) => text.includes("Supplier"));
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
    { undefinedReferences: "refuse" },
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
    { undefinedReferences: "refuse" },
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

const ROW = (id: string, text: string) => ({ id, text, kind: "paragraph" });
const idsOf = (rows: readonly { id: string }[]) => rows.map(({ id }) => id);

test("an accepted row takes the live id of the row with its text", () => {
  assert.deepEqual(
    idsOf(
      withLiveMintedIds(
        [ROW("A", "one"), ROW("X", "edited"), ROW("Y", "three")],
        [ROW("A", "one"), ROW("B", "old edited"), ROW("C", "three")],
      ),
    ),
    ["A", "B", "C"],
  );
});

test("a live row the accepted view drops keeps no accepted row", () => {
  assert.deepEqual(
    idsOf(
      withLiveMintedIds(
        [ROW("A", "one"), ROW("X", "three")],
        [ROW("A", "one"), ROW("B", "gone"), ROW("C", "three")],
      ),
    ),
    ["A", "C"],
  );
});

test("authored ids are left as they are", () => {
  assert.deepEqual(
    idsOf(
      withLiveMintedIds(
        [ROW("A", "one"), ROW("C", "three")],
        [ROW("A", "one"), ROW("B", "gone"), ROW("C", "three")],
      ),
    ),
    ["A", "C"],
  );
});

const AT = (paragraph: number) =>
  `Schema Sch_X at /w:document[1]/w:body[1]/w:p[${String(paragraph)}]/w:pPr[1]: The element has unexpected child element 'b'.`;

test("an error the source had is not new because an edit moved it", () => {
  assert.deepEqual(compareSdk([AT(846)], [AT(845)]), []);
});

test("an error the source had more of than the save is not new", () => {
  assert.deepEqual(compareSdk([AT(1), AT(2)], [AT(1)]), []);
});

test("one error more than the source had is new", () => {
  assert.deepEqual(compareSdk([AT(1)], [AT(1), AT(2)]), [AT(2)]);
});

test("another error is new", () => {
  const other = "Semantic Sem_Y at /w:numbering[1]: unexpected child 'num'.";
  assert.deepEqual(compareSdk([AT(1)], [AT(1), other]), [other]);
});

const LIVE = (...blocks: Record<string, unknown>[]) =>
  JSON.stringify({ blocks, comments: 0 });

test("a minted block id is not a change once saved", () => {
  const before = LIVE({ id: "A", idStability: "positional", text: "x" });
  const saved = LIVE({ id: "B", text: "x" });
  assert.equal(
    withoutMintedIds(saved, before),
    withoutMintedIds(before, before),
  );
});

test("a text change under a minted id is a change", () => {
  const before = LIVE({ id: "A", idStability: "positional", text: "x" });
  const saved = LIVE({ id: "A", text: "y" });
  assert.notEqual(
    withoutMintedIds(saved, before),
    withoutMintedIds(before, before),
  );
});

test("an authored id that changes is a change", () => {
  const before = LIVE({ id: "A", text: "x" });
  const saved = LIVE({ id: "B", text: "x" });
  assert.notEqual(
    withoutMintedIds(saved, before),
    withoutMintedIds(before, before),
  );
});

test("an oracle gap is recorded and the steps after it are still checked", async () => {
  const original = await mergedTableDocument();
  const reviewer = await openReviewer(original);
  const cells = reviewer
    .getContent()
    .filter((block) => block.table !== undefined);
  const first = cells.at(0);
  const last = cells.at(-1);
  assert.ok(first && last && first.id !== last.id);
  // A row in a merged table is one the oracle does not model.
  const steps = [first, last].map(({ id }) => ({
    mode: "direct" as const,
    operation: { type: "deleteTableRow" as const, blockId: id },
  }));
  const found = await runSteps(original, steps, null);
  assert.deepEqual(
    found.map(({ class: kind, operations }) => [kind, operations.length]),
    [
      ["oracle-gap", 1],
      ["oracle-gap", 2],
    ],
  );
});
