import { expect, test } from "bun:test";

import type { Comment } from "../types/content";
import { createEmptyDocument } from "../utils/createDocument";
import { withHostCommentThreads } from "./hostCommentThreads";

const comment = (id: number, parentId?: number): Comment => ({
  id,
  author: "Reviewer",
  content: [],
  ...(parentId === undefined ? {} : { parentId }),
});

test("the host's threads replace the document's, leaving the document passed in as it was", () => {
  const loaded = createEmptyDocument();
  loaded.package.document.comments = [comment(1), comment(2, 1)];
  const saved = withHostCommentThreads(
    loaded,
    [comment(1), comment(3), comment(4, 3), comment(5, 9)],
    { anchoredInEditor: null },
  );
  expect(saved.package.document.comments?.map(({ id }) => id)).toEqual([
    1, 3, 4,
  ]);
  expect(loaded.package.document.comments?.map(({ id }) => id)).toEqual([1, 2]);
});

test("an empty host list leaves the saved document without comments", () => {
  const loaded = createEmptyDocument();
  loaded.package.document.comments = [comment(1)];
  expect(
    withHostCommentThreads(loaded, [], { anchoredInEditor: null }).package
      .document.comments,
  ).toBeUndefined();
});

test("a thread the editor anchors but the saved document leaves out is left out with its replies", () => {
  const loaded = createEmptyDocument();
  loaded.package.document.content = [
    {
      type: "paragraph",
      content: [
        { type: "commentRangeStart", id: 1 },
        { type: "run", content: [{ type: "text", text: "Kept" }] },
        { type: "commentRangeEnd", id: 1 },
        { type: "commentReference", id: 1 },
      ],
    },
  ];
  const saved = withHostCommentThreads(
    loaded,
    // 7 is anchored in the editor only (a pending proposal); 9 nowhere (a
    // thread read without its anchor), which stays.
    [comment(1), comment(7), comment(8, 7), comment(9)],
    { anchoredInEditor: new Set([1, 7]) },
  );
  expect(saved.package.document.comments?.map(({ id }) => id)).toEqual([1, 9]);
});
