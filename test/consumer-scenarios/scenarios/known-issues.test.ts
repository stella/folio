/**
 * The open issues and unfiled findings, each as the smallest public-API
 * scenario that shows it. Open defects run as expected failures; after a fix,
 * the scenario stays as a passing regression without that marker.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import {
  createFolioAITextRangeHandle,
  docxToMarkdown,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "@stll/folio-core/server";
import { toMarkdown } from "@stll/folio-core/markdown";

import {
  notesDocument,
  openReviewer,
  plainDocument,
  storiesDocument,
  TEXT_BOX_TEXT,
  toArrayBuffer,
  styleNumberedDocument,
  tableDocument,
} from "../support/documents.ts";
import { saveAndReopen } from "../support/invariants.ts";
import { runFlow } from "../support/fuzz.ts";
import { expectedFailure, FINDING_SYMPTOMS, KNOWN_FAILING_FLOWS } from "../support/known-issues.ts";
import { coreBatch, MODES, type Mode, type Operation } from "../support/operations.ts";
import { applyChecked, rowsOf } from "../support/oracle.ts";
import { MARKDOWN_READ_OPTIONS } from "../support/readers.ts";
import { ENABLED_RELATIONS } from "../support/metamorphic.ts";

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

  for (const {
    seed,
    steps,
    finding,
    kind = "random",
    generation,
    relation,
  } of KNOWN_FAILING_FLOWS) {
    if (relation && !ENABLED_RELATIONS.has(relation)) continue;
    expectedFailure(
      finding,
      `the ${kind} flow with seed ${seed} (${steps} steps) does what it asked and saves it`,
      FINDING_SYMPTOMS[finding],
      async () => {
        await runFlow(seed, steps, kind, generation ? { generation } : {});
      },
    );
  }
});

test("a paragraph inserted inside a comment spanning a table reads the same before and after a save", async () => {
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
});

// Fixed findings stay as plain regressions.
describe("fixed findings", () => {
  for (const resolution of ["accept", "reject"] as const) {
    test(`${resolution}ing tracked changes in a text box and the paragraph drawing it, after a reopen, saves a package that reopens`, async () => {
      const reviewer = await openReviewer(await storiesDocument());
      const blockOf = (prefix: string) => {
        const block = reviewer.getContent().find((candidate) => candidate.text.startsWith(prefix));
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
    });
  }
});

describe("findings of the metamorphic relations (support/metamorphic.ts) and their sweeps", () => {
  const apply = (
    reviewer: Reviewer,
    mode: (typeof MODES)[number],
    operations: Record<string, unknown>[],
  ) =>
    reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: operations.map((operation, index) => ({ id: `op-${index + 1}`, ...operation })),
    } as never);
  const blockId = (reviewer: Reviewer, text: string): string => {
    const block = reviewer.getContent().find((candidate) => candidate.text === text);
    assert.ok(block, `no block "${text}"`);
    return block.id;
  };
  const reopen = async (reviewer: Reviewer) =>
    openReviewer(new Uint8Array(await reviewer.toBuffer()));

  const assertFieldKept = async (
    reviewer: Reviewer,
    text: string,
    field: "previewRuns" | "directIndentation",
  ): Promise<void> => {
    const pick = (from: Reviewer) =>
      from.getContent().find((block) => block.text === text)?.[field];
    assert.deepEqual(
      pick(await reopen(reviewer)),
      pick(reviewer),
      `the live ${field} of "${text}" is not what the saved package reads`,
    );
  };

  test("a restyled paragraph previews its new style before a save", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const text = "The Buyer pays each invoice within thirty days.";
    apply(reviewer, "direct", [
      {
        type: "setBlockParagraphProperties",
        blockId: blockId(reviewer, text),
        properties: { styleId: "Heading2" },
      },
    ]);
    await assertFieldKept(reviewer, text, "previewRuns");
  });

  test("a paragraph inserted after a bold heading previews its own direct bold off before a save", async () => {
    const reviewer = await openReviewer(await styleNumberedDocument());
    apply(reviewer, "direct", [
      { type: "insertAfterBlock", blockId: blockId(reviewer, "Definitions"), text: "Inserted." },
    ]);
    await assertFieldKept(reviewer, "Inserted.", "previewRuns");
  });

  test("a paragraph inserted with a numbered heading style reads its indentation before a save", async () => {
    const reviewer = await openReviewer(await styleNumberedDocument());
    apply(reviewer, "direct", [
      {
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "The Buyer pays on delivery."),
        text: "Inserted.",
        styleId: "Heading2",
      },
    ]);
    await assertFieldKept(reviewer, "Inserted.", "previewRuns");
    await assertFieldKept(reviewer, "Inserted.", "directIndentation");
  });

  test("a reply reads the same in toMarkdown before and after a save", async () => {
    const reviewer = await openReviewer(await plainDocument());
    apply(reviewer, "direct", [
      {
        type: "commentOnBlock",
        blockId: blockId(reviewer, "The Buyer pays each invoice within thirty days."),
        comment: { text: "Why thirty?" },
      },
    ]);
    const [comment] = reviewer.getComments();
    assert.ok(comment);
    reviewer.replyTo(comment, { text: "Market standard." });
    assert.equal(
      toMarkdown((await reopen(reviewer)).toDocument()),
      toMarkdown(reviewer.toDocument()),
      "toMarkdown reads otherwise after the save",
    );
  });

  test("deleting the last paragraph tracked and accepting it leaves what deleting it directly does", async () => {
    const texts: Record<string, string[]> = {};
    for (const mode of ["direct", "tracked-changes"] as const) {
      const reviewer = await openReviewer(await plainDocument());
      apply(reviewer, mode, [
        { type: "deleteBlock", blockId: blockId(reviewer, "Signed in two copies.") },
      ]);
      const saved = await reopen(reviewer);
      saved.acceptAll();
      texts[mode] = (await reopen(saved)).getContent().map((block) => block.text);
    }
    assert.deepEqual(
      texts["tracked-changes"],
      texts["direct"],
      "[directTracked] deleted tracked and accepted (actual) vs deleted directly (expected)",
    );
  });

  test("rejecting a tracked replacement of the last paragraph gives the document back", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const before = reviewer.getContent().map((block) => block.text);
    const last = blockId(reviewer, "Signed in two copies.");
    apply(reviewer, "tracked-changes", [
      { type: "deleteBlock", blockId: last },
      { type: "insertAfterBlock", blockId: last, text: "Inserted." },
    ]);
    const saved = await reopen(reviewer);
    saved.rejectAll();
    assert.deepEqual(
      (await reopen(saved)).getContent().map((block) => block.text),
      before,
      "[rejectAll] rejected (actual) vs the document before (expected)",
    );
  });

  test("comments on a replaced paragraph anchor alike directly and tracked-then-accepted", async () => {
    const anchors: Record<string, (string | undefined)[]> = {};
    for (const mode of ["direct", "tracked-changes"] as const) {
      const reviewer = await openReviewer(await plainDocument());
      const text = "This agreement is made between the parties named below.";
      const id = blockId(reviewer, text);
      const range = (startOffset: number, endOffset: number) =>
        createFolioAITextRangeHandle({ blockId: id, text, startOffset, endOffset });
      apply(reviewer, mode, [
        { type: "commentOnRange", range: range(0, 4), comment: { text: "a" } },
      ]);
      apply(reviewer, mode, [
        { type: "commentOnRange", range: range(31, 34), comment: { text: "b" } },
      ]);
      apply(reviewer, mode, [{ type: "replaceBlock", blockId: id, text: "New clause text." }]);
      const saved = await reopen(reviewer);
      saved.acceptAll();
      anchors[mode] = (await reopen(saved)).getComments().map((comment) => comment.anchoredText);
    }
    assert.deepEqual(
      anchors["tracked-changes"],
      anchors["direct"],
      "[directTracked] anchors tracked and accepted (actual) vs direct (expected)",
    );
  });

  test("a batch that inserts a row and deletes the row below it deletes that row", async () => {
    const reviewer = await openReviewer(await tableDocument());
    await applyChecked(
      reviewer,
      [
        { type: "deleteTableRow", blockId: blockId(reviewer, "Gadget") },
        { type: "insertTableRow", blockId: blockId(reviewer, "Widget"), position: "after" },
      ],
      "direct",
      "row insert and delete",
    );
  });
});
