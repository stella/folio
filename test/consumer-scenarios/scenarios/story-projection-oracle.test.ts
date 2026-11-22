import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FolioDocxReviewer,
} from "@stll/folio-core/server";

import {
  notesDocument,
  openFixtureBuilder,
  openReviewer,
} from "../support/documents.ts";
import {
  assertRequestedOutcome,
  capture,
  expectOperation,
  modelOf,
  resolvedState,
  rowsOf,
} from "../support/oracle.ts";

const pendingReference = async (kind: "footnote" | "endnote") => {
  // Import a pending reference deletion whose note body has no pending
  // deletion of its own. The note's independent edits are the subject of
  // these scenarios.
  const fixture = await openFixtureBuilder(await notesDocument());
  const reference = fixture
    .getContent()
    .find(
      (block) =>
        block.kind !== "diagnostic" &&
        block.structuralBoundaries?.some(
          (boundary) =>
            boundary.type === "noteReference" && boundary.noteType === kind,
        ),
    );
  assert.ok(reference, "The fixture must expose its real note reference");
  const batch = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    operations: [
      { id: "delete-reference", type: "deleteBlock", blockId: reference.id },
    ],
  } as const;
  const applied = fixture.applyDocumentOperations(
    { ...batch, mode: "tracked-changes" },
    { undefinedReferences: "refuse" },
  );
  assert.equal(applied.applied.length, 1, JSON.stringify(applied));
  assert.deepEqual(applied.skipped, []);
  const reviewer = await openReviewer(new Uint8Array(await fixture.toBuffer()));
  assert.ok(reviewer.getChanges().length > 0);
  const story = { type: kind, noteId: 1 };
  const reviewed = reviewer.readReviewedStory({
    story,
    view: "current-markup",
  });
  assert.ok(
    reviewed,
    "Pending reference deletion must leave its note editable",
  );
  assert.equal(reviewed.changes.length, 0);
  return { reviewer, story };
};

const appendTrackedNote = async ({
  reviewer,
  story,
}: Awaited<ReturnType<typeof pendingReference>>) => {
  const pre = await capture(reviewer, "tracked-changes", { story });
  const anchor = rowsOf(reviewer, story).at(0);
  assert.ok(anchor);
  const batch = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [
      {
        id: "append-note",
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Revised live note.",
      },
    ],
  } as const satisfies Parameters<typeof reviewer.applyDocumentOperations>[0];
  const receipt = reviewer.applyDocumentOperationsToStory({
    story,
    batch,
    undefinedReferences: "refuse",
  });
  assert.equal(receipt.applied.length, 1, JSON.stringify(receipt));
  return { pre, batch };
};

for (const kind of ["footnote", "endnote"] as const)
  test(`the oracle checks live ${kind} edits independently of global reference resolution`, async () => {
    const context = await pendingReference(kind);
    const { reviewer, story } = context;
    const appended = await appendTrackedNote(context);
    await assertRequestedOutcome(
      reviewer,
      appended.pre,
      { applied: appended.batch.operations },
      "append live note",
    );
    const pre = await capture(reviewer, "tracked-changes", { story });
    const block = rowsOf(reviewer, story).find(
      ({ text }) => text === "Revised live note.",
    );
    assert.ok(block);
    assert.equal(pre.resolvedStory, "present");
    assert.equal(pre.noteLifetime.type, "note");
    const range = createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 0,
      endOffset: 7,
    });
    assert.ok(range);
    const batch = {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "format-note",
          type: "formatRange",
          range,
          formatting: { italic: true },
        },
      ],
    } as const satisfies Parameters<typeof reviewer.applyDocumentOperations>[0];
    const receipt = reviewer.applyDocumentOperationsToStory({
      story,
      batch,
      undefinedReferences: "refuse",
    });
    assert.equal(receipt.applied.length, 1, JSON.stringify(receipt));
    await assertRequestedOutcome(
      reviewer,
      pre,
      { applied: batch.operations },
      "format live note",
    );
    const bytes = new Uint8Array(await reviewer.toBuffer());
    const accepted = await resolvedState({
      bytes,
      resolution: "accept",
      scope: { type: "document", story },
    });
    const rejected = await resolvedState({
      bytes,
      resolution: "reject",
      scope: { type: "document", story },
    });
    assert.equal(accepted.story, "removed");
    assert.equal(rejected.story, "present");
    assert.ok(rejected.rows.length > 0);
  });

test("a declined or incomplete scoped projection never becomes oracle evidence", async () => {
  const context = await pendingReference("footnote");
  await appendTrackedNote(context);
  const bytes = new Uint8Array(await context.reviewer.toBuffer());
  const original = bytes.slice();
  const resolve = FolioDocxReviewer.prototype.resolveReviewedStory;
  try {
    for (const result of [false, true]) {
      FolioDocxReviewer.prototype.resolveReviewedStory = () => result;
      await assert.rejects(
        resolvedState({
          bytes,
          resolution: "accept",
          scope: { type: "story", story: context.story },
        }),
        result ? /unresolved changes/u : /did not resolve/u,
      );
      assert.deepEqual(
        bytes,
        original,
        "A failed scope must not alter its input package",
      );
    }
  } finally {
    FolioDocxReviewer.prototype.resolveReviewedStory = resolve;
  }
});

test("the global lifetime backstop rejects a reference resolution that did not happen", async () => {
  const context = await pendingReference("footnote");
  const { reviewer } = context;
  const { pre, batch } = await appendTrackedNote(context);
  const accept = FolioDocxReviewer.prototype.acceptAll;
  try {
    FolioDocxReviewer.prototype.acceptAll = () => ({
      resolved: 0,
      skipped: [],
    });
    await assert.rejects(
      assertRequestedOutcome(
        reviewer,
        pre,
        { applied: batch.operations },
        "unresolved global reference",
      ),
      /reference fate/u,
    );
  } finally {
    FolioDocxReviewer.prototype.acceptAll = accept;
  }
});

test("global rejection cannot change the surviving note's locally resolved content", async () => {
  const context = await pendingReference("footnote");
  const { reviewer, story } = context;
  const { pre, batch } = await appendTrackedNote(context);
  const reject = FolioDocxReviewer.prototype.rejectAll;
  try {
    FolioDocxReviewer.prototype.rejectAll = function corruptGlobalRejection(
      this: FolioDocxReviewer,
    ) {
      const count = reject.call(this);
      const block = rowsOf(this, story).at(0);
      assert.ok(block);
      const receipt = this.applyDocumentOperationsToStory({
        story,
        undefinedReferences: "refuse",
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: [
            {
              id: "wrong-global-note",
              type: "replaceBlock",
              blockId: block.id,
              text: "Wrong global note content.",
            },
          ],
        },
      });
      assert.equal(receipt.applied.length, 1, JSON.stringify(receipt));
      return count;
    };
    await assert.rejects(
      assertRequestedOutcome(
        reviewer,
        pre,
        { applied: batch.operations },
        "corrupted global note",
      ),
      /Global reject changed surviving note content/u,
    );
  } finally {
    FolioDocxReviewer.prototype.rejectAll = reject;
  }
});

const commentedNote = async () => {
  const reviewer = await openReviewer(await notesDocument());
  const story = { type: "footnote", noteId: 1 } as const;
  const first = rowsOf(reviewer, story).at(0);
  assert.ok(first);
  const operations = [
    {
      id: "second-note-paragraph",
      type: "insertAfterBlock",
      blockId: first.id,
      text: "Another note paragraph.",
    },
    {
      id: "note-comment",
      type: "commentOnBlock",
      blockId: first.id,
      comment: { text: "Check the definition." },
    },
  ] as const;
  for (const operation of operations) {
    const receipt = reviewer.applyDocumentOperationsToStory({
      story,
      undefinedReferences: "refuse",
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [operation],
      },
    });
    assert.equal(receipt.applied.length, 1, JSON.stringify(receipt));
  }
  const root = reviewer.getComments().at(0);
  assert.ok(root);
  const reply = reviewer.replyTo(root, {
    text: "The definition is confirmed.",
  });
  assert.ok(reply);
  const replyDefinition = reviewer
    .toDocument()
    .package.document.comments?.find(({ id }) => id === reply.id);
  assert.ok(replyDefinition);
  const bytes = new Uint8Array(await reviewer.toBuffer());
  const context = await pendingReference("footnote", Promise.resolve(bytes));
  return { ...context, rootId: root.id, replyDefinition };
};

test("global note projections preserve comment paragraph positions and replies", async () => {
  const context = await commentedNote();
  const { pre, batch } = await appendTrackedNote(context);
  await assertRequestedOutcome(
    context.reviewer,
    pre,
    { applied: batch.operations },
    "commented note",
  );
  const getComments = FolioDocxReviewer.prototype.getComments;
  try {
    FolioDocxReviewer.prototype.getComments = function wrongCommentParagraph(
      this: FolioDocxReviewer,
      filter,
    ) {
      const comments = getComments.call(this, filter);
      if (
        this.getChanges().length !== 0 ||
        this.readStory(context.story) === null
      )
        return comments;
      const second = rowsOf(this, context.story).at(1);
      assert.ok(second);
      return comments.map((comment) =>
        comment.id === context.rootId
          ? Object.assign({}, comment, { blockId: second.id })
          : comment,
      );
    };
    await assert.rejects(
      assertRequestedOutcome(
        context.reviewer,
        pre,
        { applied: batch.operations },
        "misplaced note comment",
      ),
      /Global reject changed surviving note comments/u,
    );
  } finally {
    FolioDocxReviewer.prototype.getComments = getComments;
  }
});

test("global note removal cannot leave an orphan reply hidden from public threads", async () => {
  const context = await commentedNote();
  const { pre, batch } = await appendTrackedNote(context);
  const toDocument = FolioDocxReviewer.prototype.toDocument;
  try {
    FolioDocxReviewer.prototype.toDocument = function orphanReply(
      this: FolioDocxReviewer,
    ) {
      const document = toDocument.call(this);
      if (this.readStory(context.story) !== null) return document;
      return {
        ...document,
        package: {
          ...document.package,
          document: {
            ...document.package.document,
            comments: [
              ...(document.package.document.comments ?? []),
              context.replyDefinition,
            ],
          },
        },
      };
    };
    await assert.rejects(
      assertRequestedOutcome(
        context.reviewer,
        pre,
        { applied: batch.operations },
        "orphaned note reply",
      ),
      /retained comment definitions or replies/u,
    );
  } finally {
    FolioDocxReviewer.prototype.toDocument = toDocument;
  }
});

test("an actual earlier batch deletion still invalidates the later target", async () => {
  const reviewer = await openReviewer(await notesDocument());
  const rows = rowsOf(reviewer);
  const first = rows.at(0);
  assert.ok(first);
  assert.ok(rows.length > 1);
  const model = modelOf(rows);
  expectOperation(model, { type: "deleteBlock", blockId: first.id });
  assert.throws(
    () =>
      expectOperation(model, {
        type: "replaceBlock",
        blockId: first.id,
        text: "Invalid rewrite.",
      }),
    /earlier one of the batch removed/u,
  );
});
