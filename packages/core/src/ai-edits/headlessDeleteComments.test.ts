import { expect, test } from "bun:test";

import { setDocumentBodyComments } from "@stll/docx-core/model";

import { visitCommentMarkers } from "../docx/commentAnchorIndex";
import { commentTextParagraph } from "../docx/commentTextParagraph";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Comment } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const comments = [
  { id: 10, author: "Root", text: "Root thread." },
  { id: 20, author: "Reviewer", parentId: 10, text: "Selected reply." },
  { id: 30, author: "Child", parentId: 20, text: "Nested reply." },
  { id: 40, author: "Reviewer", text: "Selected root." },
  { id: 50, author: "Child", parentId: 40, text: "Dependent reply." },
  { id: 60, author: "Other", text: "Independent thread." },
] as const;

const fixture = () => {
  const document = fromMarkdown("First clause.\n\nSecond clause.");
  const body = document.package.document;
  const [first, second] = body.content;
  if (first?.type !== "paragraph" || second?.type !== "paragraph")
    throw new Error("Expected two paragraphs.");
  first.paraId = "10000001";
  second.paraId = "10000002";
  const records = comments.map((comment): Comment => {
    const paragraph = commentTextParagraph(comment.text);
    paragraph.paraId = (0x2000_0000 + comment.id).toString(16).toUpperCase();
    const record: Comment = {
      id: comment.id,
      author: comment.author,
      content: [paragraph],
    };
    if ("parentId" in comment) record.parentId = comment.parentId;
    return record;
  });
  const insertAnchors = (paragraph: typeof first, ids: readonly number[]) => {
    paragraph.content = [
      ...ids.map((id) => ({ type: "commentRangeStart" as const, id })),
      ...paragraph.content,
      ...ids.map((id) => ({ type: "commentRangeEnd" as const, id })),
      ...ids.map((id) => ({ type: "commentReference" as const, id })),
    ];
  };
  insertAnchors(first, [10, 20, 30]);
  insertAnchors(second, [40, 50, 60]);
  setDocumentBodyComments(body, records);
  return createDocx(document);
};

const assertComments = (reviewer: FolioDocxReviewer, ids: number[]) => {
  const { document } = reviewer.toDocument().package;
  expect(document.comments?.map(({ id }) => id)).toEqual(ids);
  const anchored = new Set<number>();
  visitCommentMarkers(document.content, ({ item }) => anchored.add(item.id));
  expect([...anchored].sort((left, right) => left - right)).toEqual(ids);
};

for (const session of ["legacy", "canonical"] as const) {
  test(`deletes explicit comment ids and their descendants in ${session}`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await fixture(), {
      session,
    });
    expect(reviewer.deleteComments({ ids: [50, 40, 40, 999] })).toEqual({
      deleted: [40, 50],
    });
    assertComments(reviewer, [10, 20, 30, 60]);
    expect(reviewer.deleteComments({ ids: [40, 50, 999] })).toEqual({
      deleted: [],
    });
    expect(reviewer.replyTo(50, { text: "Unavailable." })).toBeNull();
    expect(reviewer.resolveComment("40")).toBe(false);

    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    assertComments(reopened, [10, 20, 30, 60]);
    expect(reopened.getContentAsText()).toBe(reviewer.getContentAsText());
  });

  test(`deletes exactly matching authors and dependent nested replies in ${session}`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await fixture(), {
      session,
    });
    expect(reviewer.deleteComments({ author: "reviewer" })).toEqual({
      deleted: [],
    });
    expect(reviewer.deleteComments({ author: "Reviewer" })).toEqual({
      deleted: [20, 30, 40, 50],
    });
    const expected = reviewer.getComments();
    expect(expected.map(({ id, replies }) => ({ id, replies }))).toEqual([
      { id: 10, replies: [] },
      { id: 60, replies: [] },
    ]);
    assertComments(reviewer.fork(), [10, 60]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    expect(reopened.getComments()).toEqual(expected);
    assertComments(reopened, [10, 60]);
  });
}
