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
  const saved = withHostCommentThreads(loaded, [
    comment(1),
    comment(3),
    comment(4, 3),
    comment(5, 9),
  ]);
  expect(saved.package.document.comments?.map(({ id }) => id)).toEqual([
    1, 3, 4,
  ]);
  expect(loaded.package.document.comments?.map(({ id }) => id)).toEqual([1, 2]);
});

test("an empty host list leaves the saved document without comments", () => {
  const loaded = createEmptyDocument();
  loaded.package.document.comments = [comment(1)];
  expect(
    withHostCommentThreads(loaded, []).package.document.comments,
  ).toBeUndefined();
});
