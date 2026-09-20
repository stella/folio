/**
 * One traversal over the comment markers a story's blocks hold.
 *
 * Both sides of the editor round trip need it: the conversion has to know
 * which comments own a range before it decides a reference is a point comment,
 * and the save path has to know which comments still lack a reference. Two
 * walks would drift.
 */

import type { BlockContent, Paragraph, ParagraphContent } from "../types/document";
import { visitInlineContentSlots } from "./paragraphTraversal";

export type CommentMarkerSlot = {
  /** The inline-content array the marker sits in, so a caller can splice it. */
  content: ParagraphContent[];
  index: number;
  item: Extract<
    ParagraphContent,
    { type: "commentRangeStart" } | { type: "commentRangeEnd" } | { type: "commentReference" }
  >;
};

export const visitCommentMarkers = (
  blocks: readonly BlockContent[],
  visit: (slot: CommentMarkerSlot) => void,
): void => {
  const visitParagraph = (paragraph: Paragraph): void => {
    visitInlineContentSlots(paragraph, ({ content, index, item }) => {
      if (
        item.type === "commentRangeStart" ||
        item.type === "commentRangeEnd" ||
        item.type === "commentReference"
      ) {
        visit({ content, index, item });
      }
    });
  };

  const visitBlock = (block: BlockContent): void => {
    if (block.type === "paragraph") {
      visitParagraph(block);
      return;
    }
    if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          visitBlocks(cell.content);
        }
      }
      return;
    }
    if (block.type !== "blockSdt") {
      // Captured bytes hold no comment boundary folio ever read, and a
      // bookmark marker anchors none. Asking which block *does* hold children
      // keeps a new block kind out of the content control's branch.
      return;
    }
    visitBlocks(block.content);
  };

  const visitBlocks = (nested: readonly BlockContent[]): void => {
    for (const block of nested) {
      visitBlock(block);
    }
  };

  visitBlocks(blocks);
};

/**
 * Comments that open a range anywhere in these blocks.
 *
 * A `commentReference` whose id is absent has no highlight of its own — the
 * point-comment shape the model's `CommentReference` documents — and only that
 * one needs its mark anchored onto neighbouring text. Anchoring one whose range
 * exists stretches that range to wherever the reference sits, which is how a
 * comment on one word came back covering the rest of the sentence. The range
 * and the reference need not share a paragraph, so the question is asked of the
 * whole story: a reference that trails the end of a range three paragraphs back
 * is still not a point comment.
 */
export const rangedCommentIds = (blocks: readonly BlockContent[]): ReadonlySet<number> => {
  const ids = new Set<number>();
  visitCommentMarkers(blocks, ({ item }) => {
    if (item.type === "commentRangeStart") {
      ids.add(item.id);
    }
  });
  return ids;
};
