import { expect, test } from "bun:test";

import { COMMENT_PART_RELATIONSHIPS } from "../model/document";
import type { Comment, Document } from "../model/document";
import { reconcileDocumentPartRelationships } from "./comments";

const root = { id: 1, author: "Reviewer", content: [] } satisfies Comment;
// The editor projection and comment operations must allocate the same package relationships.
test.each([
  { comments: [], kinds: [] },
  { comments: [root], kinds: ["comments"] },
  {
    comments: [root, { ...root, id: 2, parentId: 1 }],
    kinds: ["comments", "commentsExtended"],
  },
  {
    comments: [{ ...root, done: true }],
    kinds: ["comments", "commentsExtended"],
  },
])(
  "comment relationship projection is complete and idempotent: %j",
  ({ comments, kinds }) => {
    const styles = { id: "rId7", type: "styles", target: "styles.xml" };
    const source = {
      package: {
        document: { content: [], comments },
        relationships: new Map([[styles.id, styles]]),
      },
    } satisfies Document;
    const projected = reconcileDocumentPartRelationships(source);
    expect(projected.package.relationships?.get(styles.id)).toEqual(styles);
    for (const [kind, part] of Object.entries(COMMENT_PART_RELATIONSHIPS)) {
      const entries = [
        ...(projected.package.relationships?.values() ?? []),
      ].filter(({ type }) => type === part.type);
      expect(entries).toHaveLength(kinds.includes(kind) ? 1 : 0);
      expect(entries.at(0)?.target).toBe(
        kinds.includes(kind) ? part.target : undefined,
      );
    }
    expect(projected.package.document.comments).toBe(
      source.package.document.comments,
    );
    expect(reconcileDocumentPartRelationships(projected)).toBe(projected);
    expect(source.package.relationships.size).toBe(1);
  },
);

test.each([
  undefined,
  { abstractNums: [], nums: [] },
  {
    abstractNums: [{ abstractNumId: 0, levels: [] }],
    nums: [{ numId: 1, abstractNumId: 0 }],
  },
])(
  "numbering relationship projection survives repeated comment reconciliation: %j",
  (numbering) => {
    const source = {
      package: { document: { content: [], comments: [root] }, numbering },
    } satisfies Document;
    const projected = reconcileDocumentPartRelationships(source);
    const entries = [
      ...(projected.package.relationships?.values() ?? []),
    ].filter(({ target }) => target === "numbering.xml");
    expect(entries).toHaveLength((numbering?.nums.length ?? 0) > 0 ? 1 : 0);
    expect(reconcileDocumentPartRelationships(projected)).toBe(projected);
    expect(source.package).not.toHaveProperty("relationships");
  },
);
