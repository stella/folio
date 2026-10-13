import { expect, test } from "bun:test";
import { COMMENT_PART_RELATIONSHIPS } from "../model/document";
import type { Comment, Document } from "../model/document";
import { reconcileCommentRelationships } from "./comments";

const root = { id: 1, author: "Reviewer", content: [] } satisfies Comment;
// The editor projection and comment operations must allocate the same package relationships.
test.each([
  { comments: [], kinds: [] },
  { comments: [root], kinds: ["comments"] },
  { comments: [root, { ...root, id: 2, parentId: 1 }], kinds: ["comments", "commentsExtended"] },
  { comments: [{ ...root, done: true }], kinds: ["comments", "commentsExtended"] },
])("comment relationship projection is complete and idempotent: %j", ({ comments, kinds }) => {
  const styles = { id: "rId7", type: "styles", target: "styles.xml" };
  const source = {
    package: { document: { content: [], comments }, relationships: new Map([[styles.id, styles]]) },
  } satisfies Document;
  const projected = reconcileCommentRelationships(source);
  expect(projected.package.relationships?.get(styles.id)).toEqual(styles);
  for (const [kind, part] of Object.entries(COMMENT_PART_RELATIONSHIPS)) {
    const entries = [...(projected.package.relationships?.values() ?? [])].filter(
      ({ type }) => type === part.type,
    );
    expect(entries).toHaveLength(kinds.includes(kind) ? 1 : 0);
    expect(entries.at(0)?.target).toBe(kinds.includes(kind) ? part.target : undefined);
  }
  expect(projected.package.document.comments).toBe(source.package.document.comments);
  expect(reconcileCommentRelationships(projected)).toBe(projected);
  expect(source.package.relationships.size).toBe(1);
});
