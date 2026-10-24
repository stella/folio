import { panic } from "better-result";
import { expect, test } from "bun:test";

import {
  openReviewer,
  paragraphsDocx,
  reviewComment,
  textRun,
} from "../__tests__/operationBatchDocuments";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";
import { withoutLostCommentThreads } from "./comment-lifecycle";

const expectSavedThreads = async (
  reviewer: Awaited<ReturnType<typeof openReviewer>>,
) => {
  const threads = (source: typeof reviewer) =>
    source.getComments().map(({ id, text, anchoredText, replies }) => ({
      id,
      text,
      anchoredText,
      replies: replies.map(({ id: replyId, text: replyText }) => ({
        id: replyId,
        text: replyText,
      })),
    }));
  const reopened = await openReviewer(await reviewer.toBuffer());
  expect(threads(reopened)).toEqual(threads(reviewer));
  const comments = reopened.toDocument().package.document.comments ?? [];
  const ids = new Set(comments.map(({ id }) => id));
  expect(ids.size).toBe(comments.length);
  for (const { parentId } of comments) {
    if (parentId !== undefined) expect(ids.has(parentId)).toBe(true);
  }
  return comments;
};

test("retiring a replaced comment anchor also retires its newly created reply", async () => {
  const reviewer = await openReviewer(
    await paragraphsDocx(
      [
        [
          { type: "commentRangeStart", id: 0 },
          textRun("Paragraph 0 of the agreement."),
          { type: "commentRangeEnd", id: 0 },
          { type: "commentReference", id: 0 },
        ],
      ],
      [reviewComment(0, "A")],
    ),
  );
  await expectSavedThreads(reviewer);
  const block = reviewer.snapshot().blocks.at(0);
  if (!block) panic("Expected the commented paragraph");
  const applied = reviewer.applyOperations(
    [
      {
        id: "replace",
        type: "replaceInBlock",
        blockId: block.id,
        find: block.text,
        replace: "Revised.",
      },
    ],
    { undefinedReferences: "refuse", mode: "tracked-changes" },
  );
  expect(applied.skipped).toEqual([]);
  await expectSavedThreads(reviewer);
  const change = reviewer.getChanges().at(0);
  if (!change) panic("Expected a tracked replacement change");
  expect(reviewer.acceptChange(change)).toBe(true);
  await expectSavedThreads(reviewer);
  const comment = reviewer.getComments().at(0);
  if (!comment) panic("Expected the surviving comment thread");
  reviewer.replyTo(comment, { text: "Session reply." });
  await expectSavedThreads(reviewer);
  expect(
    reviewer
      .getComments()
      .at(0)
      ?.replies.map(({ text }) => text),
  ).toEqual(["Session reply."]);
  expect(reviewer.rejectAll()).toBeGreaterThan(0);
  await expectSavedThreads(reviewer);
  const reopened = await openReviewer(await reviewer.toBuffer());
  expect(reviewer.getComments()).toEqual([]);
  expect(reopened.getComments()).toEqual([]);
});

const anchoredParagraph = (id: number, text: string): Paragraph => ({
  type: "paragraph",
  paraId: (0x20000000 + id).toString(16),
  content: [
    { type: "commentRangeStart", id },
    textRun(text),
    { type: "commentRangeEnd", id },
    { type: "commentReference", id },
  ],
});

test("retiring a body thread preserves unrelated body, note and source-unanchored threads", async () => {
  const target = anchoredParagraph(0, "Original.");
  const other = anchoredParagraph(20, "Other story.");
  const survivor = anchoredParagraph(7, "Survivor.");
  survivor.content.push({
    type: "run",
    content: [{ type: "footnoteRef", id: 2 }],
  });
  const reviewer = await openReviewer(
    await createDocx({
      package: {
        document: {
          content: [target, survivor],
          comments: [
            reviewComment(0),
            { ...reviewComment(1), parentId: 0 },
            reviewComment(7),
            reviewComment(20),
            reviewComment(30),
            { ...reviewComment(31), parentId: 30 },
          ],
        },
        footnotes: [{ type: "footnote", id: 2, content: [other] }],
      },
    }),
  );
  await expectSavedThreads(reviewer);
  const block = reviewer.snapshot().blocks.at(0);
  if (!block) panic("Expected the commented story block");
  const result = reviewer.applyOperations(
    [
      {
        id: "replace",
        type: "replaceInBlock",
        blockId: block.id,
        find: block.text,
        replace: "Revised.",
      },
    ],
    {
      undefinedReferences: "refuse",
      mode: "tracked-changes",
    },
  );
  expect(result.skipped).toEqual([]);
  await expectSavedThreads(reviewer);
  const change = reviewer.getChanges().at(0);
  if (!change) panic("Expected a tracked story change");
  expect(reviewer.acceptChange(change)).toBe(true);
  await expectSavedThreads(reviewer);
  expect(reviewer.replyTo(0, { text: "Session reply." })).not.toBeNull();
  await expectSavedThreads(reviewer);
  expect(reviewer.rejectAll()).toBeGreaterThan(0);
  const comments = await expectSavedThreads(reviewer);
  expect(
    comments.map(({ id }) => id).sort((left, right) => left - right),
  ).toEqual([7, 20, 30, 31]);
});

const ANCHOR_TRANSITIONS = {
  retired: { before: [0, 100], now: [100], retainTarget: false },
  retained: { before: [0, 100], now: [0, 100], retainTarget: true },
  sourceUnanchored: { before: [100], now: [100], retainTarget: true },
} as const;

test.each(Object.entries(ANCHOR_TRANSITIONS))(
  "canonical thread ownership cascades nested replies only for %s roots",
  (_state, { before, now, retainTarget }) => {
    for (const depth of [1, 2, 3, 4, 5]) {
      const target = [reviewComment(0)];
      for (let id = 1; id <= depth; id++)
        target.push({ ...reviewComment(id), parentId: id - 1 });
      const unrelated = [
        reviewComment(100),
        { ...reviewComment(101), parentId: 100 },
        reviewComment(200),
      ];
      const comments = [...target, ...unrelated];
      const original = JSON.stringify(comments);
      const kept = withoutLostCommentThreads(comments, {
        anchoredBefore: new Set(before),
        anchoredNow: new Set(now),
      });
      expect(kept).toEqual(retainTarget ? comments : unrelated);
      expect(JSON.stringify(comments)).toBe(original);
    }
  },
);
