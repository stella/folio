/**
 * The open issues and unfiled findings, each as the smallest public-API
 * scenario that shows it. Every one runs as an expected failure: when a fix
 * lands, its scenario passes, the expected failure fails, and the marker comes
 * off so the scenario guards the fix.
 */

import assert from "node:assert/strict";
import { describe } from "node:test";

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";
import {
  createFolioAITextRangeHandle,
  docxToMarkdown,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "@stll/folio-core/server";

import {
  notesDocument,
  openReviewer,
  plainDocument,
  storiesDocument,
  TEXT_BOX_TEXT,
  toArrayBuffer,
  unusedNumberingDocument,
} from "../support/documents.ts";
import { saveAndReopen } from "../support/invariants.ts";
import { runFlow } from "../support/fuzz.ts";
import { expectedFailure, FINDING_SYMPTOMS, KNOWN_FAILING_FLOWS } from "../support/known-issues.ts";
import { coreBatch, MODES, type Mode, type Operation } from "../support/operations.ts";
import { rowsOf } from "../support/oracle.ts";
import { MARKDOWN_READ_OPTIONS } from "../support/readers.ts";

const MISSING_NUMBERING = /Numbering definition \d+ is missing/u;

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

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Story = Parameters<Reviewer["snapshotStory"]>[0];

const applyTo = (reviewer: Reviewer, story: Story, operation: Operation, mode: Mode) => {
  const batch = coreBatch([operation], mode) as never;
  const result =
    story.type === "main"
      ? reviewer.applyDocumentOperations(batch)
      : reviewer.applyDocumentOperationsToStory({ story, batch });
  assert.equal(result.applied.length, 1, JSON.stringify(result.issues));
};

describe("findings not yet filed", () => {
  expectedFailure(
    "SUGGESTED_NOTE_EDIT_SAVE_THROWS",
    "a footnote suggested twice around a save, with body suggestions, saves",
    FINDING_SYMPTOMS.SUGGESTED_NOTE_EDIT_SAVE_THROWS,
    async () => {
      const footnote = { type: "footnote", noteId: 1 } as const;
      let reviewer = await openReviewer(await notesDocument());
      const note = () => rowsOf(reviewer, footnote)[0] as { id: string; text: string };
      const lastBody = () => reviewer.getContent().at(-1) as { id: string };
      const { id, text } = note();
      const range = createFolioAITextRangeHandle({
        blockId: id,
        text,
        startOffset: 0,
        endOffset: 7,
      });
      applyTo(
        reviewer,
        footnote,
        { type: "formatRange", range, formatting: { italic: true } },
        "suggested",
      );
      applyTo(
        reviewer,
        { type: "main" },
        { type: "insertAfterBlock", blockId: lastBody().id, text: "One." },
        "suggested",
      );
      // Suggestions stay out of the package: this saves the document before them.
      reviewer = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
      applyTo(
        reviewer,
        footnote,
        { type: "replaceBlock", blockId: note().id, text: "Rewritten note." },
        "suggested",
      );
      applyTo(
        reviewer,
        { type: "main" },
        { type: "insertBeforeBlock", blockId: lastBody().id, text: "Two." },
        "suggested",
      );
      await reviewer.toBuffer();
    },
  );

  for (const resolution of ["accept", "reject"] as const) {
    expectedFailure(
      "TEXT_BOX_RESOLVE_MALFORMED_XML",
      `${resolution}ing tracked changes in a text box and the paragraph drawing it, after a reopen, saves a package that reopens`,
      FINDING_SYMPTOMS.TEXT_BOX_RESOLVE_MALFORMED_XML,
      async () => {
        const reviewer = await openReviewer(await storiesDocument());
        const blockOf = (prefix: string) => {
          const block = reviewer
            .getContent()
            .find((candidate) => candidate.text.startsWith(prefix));
          assert.ok(block, prefix);
          return block.id;
        };
        const edit = (prefix: string, find: string) =>
          applyTo(
            reviewer,
            { type: "main" },
            { type: "replaceInBlock", blockId: blockOf(prefix), find, replace: "amended" },
            "tracked-changes",
          );
        edit("The box beside", "beside");
        edit(TEXT_BOX_TEXT, "Boxed");
        const reopened = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
        if (resolution === "accept") reopened.acceptAll();
        else reopened.rejectAll();
        await openReviewer(new Uint8Array(await reopened.toBuffer()));
      },
    );
  }

  expectedFailure(
    "TRACKED_MERGE_INTO_DELETED_BLOCK",
    "a tracked batch that merges a block into one it deletes is refused, as in direct mode",
    FINDING_SYMPTOMS.TRACKED_MERGE_INTO_DELETED_BLOCK,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const idOf = (prefix: string) => {
        const block = reviewer.getContent().find((candidate) => candidate.text.startsWith(prefix));
        assert.ok(block, prefix);
        return block.id;
      };
      const result = reviewer.applyDocumentOperations(
        coreBatch(
          [
            { type: "deleteBlock", blockId: idOf("The Supplier") },
            { type: "mergeBlockWithNext", blockId: idOf("This agreement"), separator: " " },
          ],
          "tracked-changes",
        ) as never,
      );
      assert.deepEqual(
        result.issues.map((issue) => issue.code),
        ["overlappingOperation"],
        "applied a merge into a block the batch deletes",
      );
    },
  );

  expectedFailure(
    "INSERT_AFTER_PENDING_MERGE",
    "a paragraph inserted after a block with a pending tracked merge stays its own paragraph once accepted",
    FINDING_SYMPTOMS.INSERT_AFTER_PENDING_MERGE,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const buyer = reviewer.getContent().find((block) => block.text.startsWith("The Buyer"));
      assert.ok(buyer);
      for (const operation of [
        { type: "mergeBlockWithNext", blockId: buyer.id, separator: " " },
        { type: "insertAfterBlock", blockId: buyer.id, text: "New clause." },
      ]) {
        applyTo(reviewer, { type: "main" }, operation, "tracked-changes");
      }
      const saved = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
      saved.acceptAll();
      const accepted = (await openReviewer(new Uint8Array(await saved.toBuffer())))
        .getContent()
        .map((block) => block.text);
      assert.ok(
        accepted.includes("New clause."),
        `accepting glues the inserted paragraph onto the merged one: ${JSON.stringify(accepted)}`,
      );
    },
  );

  expectedFailure(
    "MARKDOWN_DROPS_TEXT_BOX",
    "docxToMarkdown reads a text box's paragraph that getContent() lists",
    FINDING_SYMPTOMS.MARKDOWN_DROPS_TEXT_BOX,
    async () => {
      const bytes = await storiesDocument();
      const reviewer = await openReviewer(bytes);
      assert.ok(reviewer.getContent().some((block) => block.text === TEXT_BOX_TEXT));
      const markdown = await docxToMarkdown(toArrayBuffer(bytes), MARKDOWN_READ_OPTIONS);
      assert.ok(
        markdown.includes(TEXT_BOX_TEXT),
        `docxToMarkdown writes no text-box paragraph:\n${markdown}`,
      );
    },
  );

  expectedFailure(
    "LIVE_COMMENT_ANCHOR_STALE",
    "a paragraph inserted inside a comment spanning a table reads the same before and after a save",
    FINDING_SYMPTOMS.LIVE_COMMENT_ANCHOR_STALE,
    async () => {
      const reviewer = await openReviewer(await storiesDocument());
      const start = reviewer
        .getContent()
        .find((block) => block.text === "The schedule below is binding.");
      assert.ok(start);
      applyTo(
        reviewer,
        { type: "main" },
        { type: "insertAfterBlock", blockId: start.id, text: "New clause." },
        "direct",
      );
      await saveAndReopen(reviewer, "insert inside a comment range");
    },
  );

  for (const { seed, steps, finding, kind = "random", generation } of KNOWN_FAILING_FLOWS) {
    expectedFailure(
      finding,
      `the ${kind} flow with seed ${seed} (${steps} steps) does what it asked and saves it`,
      FINDING_SYMPTOMS[finding],
      () => runFlow(seed, steps, kind, generation ? { generation } : {}),
    );
  }
});
