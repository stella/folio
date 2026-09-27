/**
 * `FolioDocxReviewer.applyDocumentOperations` over every fixture, in every
 * mode, one operation type at a time: each batch applies or refuses with an
 * issue, the result saves and reopens, every reader agrees, and resolving the
 * changes lands where it should (reject-all: the document before; accept-all:
 * the document a reader saw).
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "@stll/folio-core/server";

import { FIXTURE_NAMES, FIXTURES, openReviewer, plainDocument } from "../support/documents.ts";
import { assertHealthy, saveAndReopen, visibleState } from "../support/invariants.ts";
import { type Block, coreBatch, GENERATORS, MODES, supports } from "../support/operations.ts";
import { createRandom } from "../support/random.ts";
import { expectedFailure, KNOWN_FAILING_OPERATION_RUNS } from "../support/known-issues.ts";
import { resolvedText, settledText } from "../support/review.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

const blocksOf = (reviewer: Reviewer): Block[] => reviewer.getContent() as Block[];

describe("applyDocumentOperations", () => {
  for (const name of FIXTURE_NAMES) {
    for (const mode of MODES) {
      const known = KNOWN_FAILING_OPERATION_RUNS.find(
        (run) => run.fixture === name && run.mode === mode,
      );
      const title = `${name} / ${mode}: every operation type applies or refuses, and the result saves`;
      const register = (body: () => Promise<void>) =>
        known
          ? expectedFailure(known.finding, title, /out of range|nodeSize/u, body)
          : test(title, body);
      register(async () => {
        const before = await FIXTURES[name]();
        const reviewer = await openReviewer(before);
        // `"suggested"` edits stay out of the saved package until accepted.
        const persisted = mode === "suggested" ? { persisted: visibleState(reviewer) } : {};
        const random = createRandom(name.length * 31 + mode.length);
        const log: string[] = [];
        for (const type of Object.keys(GENERATORS)) {
          if (!supports(type, mode)) continue;
          const operation = GENERATORS[type]?.(blocksOf(reviewer), random);
          if (!operation) continue;
          const result = reviewer.applyDocumentOperations(coreBatch([operation], mode) as never);
          const outcome =
            result.applied.length > 0
              ? "applied"
              : `refused ${result.issues.map((issue) => issue.code).join(",")}`;
          log.push(`${type}: ${outcome}`);
          assert.ok(
            result.applied.length + result.skipped.length === 1,
            `${type} neither applied nor refused: ${JSON.stringify(result)}`,
          );
          if (result.skipped.length > 0) {
            assert.ok(result.issues.length > 0, `${type} was skipped without an issue`);
          }
          await assertHealthy(reviewer, `${name} / ${mode} after ${log.join(" → ")}`, persisted);
        }
        assert.ok(
          log.some((entry) => entry.endsWith("applied")),
          `nothing applied: ${log.join("; ")}`,
        );

        if (mode === "suggested") {
          // Accepted, the suggestions are ordinary content and save as such.
          reviewer.acceptAll();
        }
        const { bytes: after } = await assertHealthy(reviewer, `${name} / ${mode} final`);
        // Rejecting gives the document back, block for block. Accepting lands
        // on the words a reader saw; block boundaries are left out there, as
        // a reader shows a pending join or split as the blocks it has now,
        // and so is whitespace: a merge's separator belongs to neither block
        // until the join is accepted.
        const words = (blocks: string[]) =>
          blocks
            .map((block) => block.replace(/^\w+: /u, ""))
            .join("")
            .replace(/\s+/gu, "");
        if (mode === "tracked-changes") {
          assert.deepEqual(
            await resolvedText(after, "reject"),
            await resolvedText(before, "reject"),
            `${name} / ${mode}: rejecting every change does not give the document back`,
          );
        }
        assert.equal(
          words(await resolvedText(after, "accept")),
          words(settledText(await openReviewer(after))),
          `${name} / ${mode}: accepting every change does not give what readers showed`,
        );
      });
    }
  }
});

describe("a batch that splits a block and deletes it", () => {
  test("refuses the deletion, whose block the split already claims, and loses no word", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const text = "The Buyer pays each invoice within thirty days.";
    const target = reviewer.getContent().find((block) => block.text === text);
    assert.ok(target);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        { id: "split", type: "splitBlock", blockId: target.id, offset: 20 },
        { id: "delete", type: "deleteBlock", blockId: target.id },
      ],
    });
    assert.deepEqual(
      result.applied.map(({ id }) => id),
      ["split"],
    );
    assert.deepEqual(
      result.issues.map(({ operationId, code }) => ({ operationId, code })),
      [{ operationId: "delete", code: "overlappingOperation" }],
    );
    reviewer.acceptAll();
    const texts = reviewer.getContent().map((block) => block.text);
    assert.ok(
      texts.includes(text.slice(0, 20)) && texts.includes(text.slice(20)),
      `the split block's two halves: ${JSON.stringify(texts)}`,
    );
  });
});

describe("resolving tracked edits that build on pending ones", () => {
  const SUPPLIER = "The Supplier delivers the goods on time and in good order.";
  const tracked = async () => {
    const reviewer = await openReviewer(await plainDocument());
    const apply = (operation: Record<string, unknown>) =>
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [{ id: "1", ...operation }],
      } as never);
    const block = (prefix: string) => {
      const found = reviewer.getContent().find(({ text }) => text.startsWith(prefix));
      assert.ok(found, `no block starts with "${prefix}"`);
      return found;
    };
    return { reviewer, apply, block };
  };
  const originalText = async () =>
    (await openReviewer(await plainDocument())).getContent().map(({ text }) => text);

  test("rejectAll undoes a merge of a split's second half into an inserted paragraph", async () => {
    const { reviewer, apply, block } = await tracked();
    apply({ type: "insertAfterBlock", blockId: block("Signed").id, text: "Inserted clause." });
    apply({ type: "splitBlock", blockId: block("Signed").id, offset: "Signed in ".length });
    apply({ type: "mergeBlockWithNext", blockId: block("two copies").id, separator: " " });
    reviewer.rejectAll();
    assert.deepEqual(
      reviewer.getContent().map(({ text }) => text),
      await originalText(),
    );
  });

  test("rejecting a split with a table inserted between its halves joins them again", async () => {
    const { reviewer, apply, block } = await tracked();
    const target = block("The Supplier");
    apply({ type: "splitBlock", blockId: target.id, offset: SUPPLIER.indexOf("good order") });
    apply({ type: "insertTable", blockId: target.id, rows: [["Term", "Value"]] });
    reviewer.rejectAll();
    assert.deepEqual(
      reviewer.getContent().map(({ text }) => text),
      await originalText(),
    );
  });

  test("a comment's anchored text reads the same before and after a save when its text is replaced", async () => {
    const { reviewer, apply, block } = await tracked();
    const target = block("The Supplier");
    const start = SUPPLIER.indexOf("good order");
    const range = createFolioAITextRangeHandle({
      blockId: target.id,
      text: SUPPLIER,
      startOffset: start,
      endOffset: start + "good".length,
    });
    assert.ok(range);
    apply({ type: "commentOnRange", range, comment: { text: "Which standard?" } });
    apply({ type: "replaceBlock", blockId: target.id, text: "The Supplier delivers promptly." });
    const before = reviewer.getComments().map((comment) => comment.anchoredText);
    // The removed word, and the new one that takes its place.
    assert.deepEqual(before, ["goodpromptly"]);
    const { reopened } = await saveAndReopen(reviewer, "comment anchor");
    assert.deepEqual(
      reopened.getComments().map((comment) => comment.anchoredText),
      before,
      "the anchored text changed across the save",
    );
  });
});
