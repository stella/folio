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
