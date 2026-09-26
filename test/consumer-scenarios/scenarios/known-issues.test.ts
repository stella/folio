/**
 * The open issues and unfiled findings, each as the smallest public-API
 * scenario that shows it. Every one runs as an expected failure: when a fix
 * lands, its scenario passes, the expected failure fails, and the marker comes
 * off so the scenario guards the fix.
 */

import assert from "node:assert/strict";
import { describe } from "node:test";

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";
import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  generateRedlineDocx,
  paragraph,
  run,
} from "@stll/folio-core/server";

import {
  directNumberedDocument,
  listDocument,
  openReviewer,
  packDocument,
  plainDocument,
  toArrayBuffer,
  unusedNumberingDocument,
} from "../support/documents.ts";
import { assertReadersAgree, saveAndReopen } from "../support/invariants.ts";
import { runFlow } from "../support/fuzz.ts";
import { expectedFailure, KNOWN_FAILING_FLOWS } from "../support/known-issues.ts";
import { MODES } from "../support/operations.ts";

const MISSING_NUMBERING = /Numbering definition \d+ is missing/u;

const labelsOf = async (bytes: Uint8Array): Promise<string[]> =>
  (await openReviewer(bytes))
    .getContent()
    .map((block) => `${block.displayLabel ?? "·"} ${block.text}`);

describe("#1103: operations that name a numbering instance the package does not define", () => {
  for (const mode of MODES) {
    expectedFailure(
      1103,
      `insertAfterBlock with an undefined numId (${mode}) is refused or saves`,
      MISSING_NUMBERING,
      async () => {
        const reviewer = await openReviewer(await unusedNumberingDocument());
        const anchor = reviewer
          .getContent()
          .find((block) => block.text === "Signed in two copies.");
        assert.ok(anchor);
        reviewer.applyDocumentOperations({
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "1",
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "An inserted clause.",
              numbering: { numId: 1, level: 0 },
            },
          ],
        });
        if (mode === "suggested") {
          // A suggestion reaches the package once accepted.
          reviewer.acceptAll();
        }
        await saveAndReopen(reviewer, `#1103 ${mode}`);
      },
    );
  }

  expectedFailure(
    1103,
    "setBlockParagraphProperties with an undefined numId is refused or saves",
    MISSING_NUMBERING,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const target = reviewer.getContent().find((block) => block.text === "Signed in two copies.");
      assert.ok(target);
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "1",
            type: "setBlockParagraphProperties",
            blockId: target.id,
            properties: { numbering: { numId: 3, level: 0 } },
          },
        ],
      });
      await saveAndReopen(reviewer, "#1103 setBlockParagraphProperties");
    },
  );

  expectedFailure(
    1103,
    "suggest_changes with an undefined numId answers with an issue, and the document saves",
    MISSING_NUMBERING,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const anchor = reviewer.getContent().find((block) => block.text === "Signed in two copies.");
      assert.ok(anchor);
      executeFolioToolCallUntyped(
        "suggest_changes",
        {
          operations: [
            {
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "An inserted clause.",
              numbering: { numId: 1, level: 0 },
            },
          ],
        },
        createReviewerBridge(reviewer, { mode: "tracked-changes" }),
        {},
      );
      await saveAndReopen(reviewer, "#1103 suggest_changes");
    },
  );
});

describe("findings not yet filed", () => {
  expectedFailure(
    "STALE_LIST_LABELS",
    "an item inserted into a list reads its own number, and the items after it renumber, before a save",
    /labels before the save/u,
    async () => {
      const reviewer = await openReviewer(await listDocument());
      const anchor = reviewer.getContent().find((block) => block.text === "Deposit on signature");
      assert.ok(anchor);
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          { id: "1", type: "insertAfterBlock", blockId: anchor.id, text: "Interim payment" },
        ],
      });
      const live = reviewer
        .getContent()
        .filter((block) => block.listReference?.numId === anchor.listReference?.numId)
        .map((block) => `${block.displayLabel} ${block.text}`);
      assert.deepEqual(
        live,
        [
          "1. Deposit on signature",
          "2. Interim payment",
          "3. Balance on delivery",
          "4. Retention after inspection",
        ],
        "labels before the save are stale",
      );
    },
  );

  expectedFailure(
    "COMMENT_ANCHOR_DRIFT",
    "a comment's anchored text reads the same before and after a save when its text is replaced",
    /anchored text/u,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const text = "The Supplier delivers the goods on time and in good order.";
      const target = reviewer.getContent().find((block) => block.text === text);
      assert.ok(target);
      const start = text.indexOf("good order");
      const range = createFolioAITextRangeHandle({
        blockId: target.id,
        text,
        startOffset: start,
        endOffset: start + "good".length,
      });
      assert.ok(range);
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          { id: "comment", type: "commentOnRange", range, comment: { text: "Which standard?" } },
        ],
      });
      // Before the save the comment covers "good" and the replacement text;
      // after it, "good order." and the replacement text.
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "replace",
            type: "replaceBlock",
            blockId: target.id,
            text: "The Supplier delivers promptly.",
          },
        ],
      });
      const before = reviewer.getComments().map((comment) => comment.anchoredText);
      const { reopened } = await saveAndReopen(reviewer, "comment anchor");
      assert.deepEqual(
        reopened.getComments().map((comment) => comment.anchoredText),
        before,
        "the anchored text changed across the save",
      );
    },
  );

  expectedFailure(
    "COMPARE_INSERTED_LIST_ITEMS",
    "accepting a redline keeps an inserted bullet a bullet",
    /inserted list item/u,
    async () => {
      const before = await packDocument(fromMarkdown("Intro.\n\nOutro."));
      const after = await packDocument(fromMarkdown("Intro.\n\n- new bullet\n\nOutro."));
      const redline = await generateRedlineDocx(toArrayBuffer(before), toArrayBuffer(after));
      const reviewer = await openReviewer(new Uint8Array(redline.buffer));
      reviewer.acceptAll();
      assert.deepEqual(
        await labelsOf(new Uint8Array(await reviewer.toBuffer())),
        ["· Intro.", "• new bullet", "· Outro."],
        "the inserted list item lost its bullet",
      );
    },
  );

  expectedFailure(
    "REJECT_SPLIT_AROUND_INSERTED_TABLE",
    "rejecting a split with a table inserted between its halves joins them again",
    /split/u,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const text = "The Supplier delivers the goods on time and in good order.";
      const target = reviewer.getContent().find((block) => block.text === text);
      assert.ok(target);
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "split",
            type: "splitBlock",
            blockId: target.id,
            offset: text.indexOf("good order"),
          },
        ],
      });
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          { id: "table", type: "insertTable", blockId: target.id, rows: [["Term", "Value"]] },
        ],
      });
      reviewer.rejectAll();
      assert.ok(
        reviewer.getContent().some((block) => block.text === text),
        "the split is still there after rejecting every change",
      );
    },
  );

  expectedFailure(
    "BATCH_SPLIT_THEN_DELETE",
    "a batch that splits a block and deletes it removes all of its text (or refuses)",
    /survives/u,
    async () => {
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
      if (result.applied.length < 2) return;
      reviewer.acceptAll();
      const left = reviewer
        .getContent()
        .map((block) => block.text)
        .filter((blockText) => blockText.length === 0 || text.includes(blockText));
      assert.deepEqual(left, [], "text of the deleted block survives");
    },
  );

  expectedFailure(
    "NOTE_REFERENCE_TEXT",
    "a note reference reads as the number the page shows, not its w:id",
    /note reference/u,
    async () => {
      const document = fromMarkdown("Intro.");
      document.package.footnotes = [
        { type: "footnote", id: 7, content: [paragraph("First note.")] },
      ];
      document.package.document.content.push(
        paragraph([run("Referenced."), { type: "run", content: [{ type: "footnoteRef", id: 7 }] }]),
      );
      const reviewer = await openReviewer(await packDocument(document));
      assert.equal(
        reviewer.getContent().at(-1)?.text,
        "Referenced.1",
        "the note reference reads as its id",
      );
    },
  );

  expectedFailure(
    "REJECT_ALL_JOIN_INTO_INSERTED_PARAGRAPH",
    "rejectAll undoes a merge of a split's second half into an inserted paragraph",
    /out of range|nodeSize/u,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const apply = (operation: Record<string, unknown>) =>
        reviewer.applyDocumentOperations({
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [{ id: "1", ...operation }],
        } as never);
      const last = () => {
        const block = reviewer.getContent().find(({ text }) => text.startsWith("Signed"));
        assert.ok(block);
        return block;
      };
      apply({ type: "insertAfterBlock", blockId: last().id, text: "Inserted clause." });
      apply({ type: "splitBlock", blockId: last().id, offset: "Signed in ".length });
      const blocks = reviewer.getContent();
      const secondHalf = blocks[blocks.findIndex(({ id }) => id === last().id) + 1];
      assert.ok(secondHalf);
      apply({ type: "mergeBlockWithNext", blockId: secondHalf.id, separator: " " });
      reviewer.rejectAll();
      assert.deepEqual(
        reviewer.getContent().map(({ text }) => text),
        (await openReviewer(await plainDocument())).getContent().map(({ text }) => text),
      );
    },
  );

  expectedFailure(
    "UNMARKED_LIST_ITEM_KIND",
    "a paragraph numbered at a level its instance does not define reads alike everywhere",
    /docxToMarkdown vs getContent/u,
    async () => {
      const reviewer = await openReviewer(await directNumberedDocument());
      const anchor = reviewer.getContent().find(({ text }) => text === "Unnumbered body text.");
      assert.ok(anchor);
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          {
            id: "1",
            type: "insertAfterBlock",
            blockId: anchor.id,
            text: "Level eight.",
            numbering: { numId: 7, level: 8 },
          },
        ],
      });
      await assertReadersAgree(new Uint8Array(await reviewer.toBuffer()), "undefined level", {
        strict: true,
      });
    },
  );

  for (const { seed, steps, finding } of KNOWN_FAILING_FLOWS) {
    expectedFailure(
      finding,
      `the fuzz flow with seed ${seed} (${steps} steps) saves what the reviewer shows`,
      /reopened package shows something else/u,
      () => runFlow(seed, steps),
    );
  }
});
