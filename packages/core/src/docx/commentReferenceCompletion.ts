/**
 * Every comment a story anchors keeps exactly one `w:commentReference`.
 *
 * The reference is the run Word paints the visible comment mark with, and its
 * position among the range ends around it is what a reader sees as the marks'
 * order. That position is authored data, so the editor carries it as a node
 * and `fromProseDoc` writes it back where it sat. The serializer therefore
 * writes what the model says and never invents a reference of its own: a
 * per-paragraph guess put one after every range end, which regrouped the marks
 * at a shared boundary and painted a comment spanning three paragraphs three
 * times.
 *
 * A model can still arrive with a range and no reference at all — a package
 * written by a library that omitted it — and Word shows no mark for such a
 * comment. This closes that gap once, for the whole story, by placing the
 * missing reference after the comment's last range end.
 *
 * The ProseMirror conversion calls it, and that conversion owns every story an
 * edit can produce: `fromProseDoc` for the body, `proseDocToBlocks` for a
 * header, footer or note. A headless edit arrives through the same conversion,
 * and markdown has no comment syntax to import one through. The save path
 * deliberately does not complete a reference of its own: it keeps the
 * paragraphs an edit did not touch byte-for-byte, so one minted there would
 * sit in the model and never reach the package.
 */

import type { BlockContent, ParagraphContent } from "../types/document";
import { visitCommentMarkers } from "./commentAnchorIndex";

type LastRangeEnd = {
  content: ParagraphContent[];
  index: number;
};

/**
 * Give every comment that a `commentRangeEnd` closes but no `commentReference`
 * names a reference after its last end. Answers how many were added.
 */
export const completeCommentReferences = (blocks: readonly BlockContent[]): number => {
  const lastRangeEnds = new Map<number, LastRangeEnd>();
  const referencedIds = new Set<number>();

  visitCommentMarkers(blocks, ({ content, index, item }) => {
    if (item.type === "commentReference") {
      referencedIds.add(item.id);
      return;
    }
    if (item.type === "commentRangeEnd") {
      lastRangeEnds.set(item.id, { content, index });
    }
  });

  // Inserting shifts every later index in the same array, so the insertions
  // are applied per array from the back.
  const insertionsByContent = new Map<ParagraphContent[], { index: number; id: number }[]>();
  for (const [id, { content, index }] of lastRangeEnds) {
    if (referencedIds.has(id)) {
      continue;
    }
    const pending = insertionsByContent.get(content);
    if (pending) {
      pending.push({ index, id });
      continue;
    }
    insertionsByContent.set(content, [{ index, id }]);
  }

  let added = 0;
  for (const [content, pending] of insertionsByContent) {
    for (const { index, id } of pending.toSorted(
      (first, second) => second.index - first.index || second.id - first.id,
    )) {
      content.splice(index + 1, 0, { type: "commentReference", id });
      added += 1;
    }
  }
  return added;
};
