/**
 * A comment owns one range in a story, however many paragraphs it crosses.
 *
 * `w:commentRangeStart` and `w:commentRangeEnd` are story positions, not
 * paragraph ones: a comment on three paragraphs is one start in the first and
 * one end in the last, with nothing in between. The editor carries the range as
 * a mark, and the save path reads each paragraph on its own, so a mark that
 * crosses a paragraph boundary closes at the boundary and reopens after it.
 * The document then holds one range per paragraph where the author wrote one,
 * which moves the range's end to the end of its first paragraph and repaints
 * the comment once per paragraph.
 *
 * This joins them once, for the whole story: the first start and the last end
 * of an id are the range the author wrote, and every marker between them is a
 * boundary artefact. References are untouched — where a reference sits is
 * authored data (see `commentReferenceCompletion`).
 */

import type { BlockContent } from "../types/document";
import type { CommentMarkerSlot } from "./commentAnchorIndex";
import { visitCommentMarkers } from "./commentAnchorIndex";
import { InlineContentRemovals } from "./inlineContentRemovals";

type MarkerSlot = CommentMarkerSlot;

/**
 * Collapse each comment's range markers to its first start and its last end.
 * Answers how many markers were dropped.
 *
 * A start or an end between blocks counts as one of the range's markers, so a
 * range that opens beside a content control keeps that start and loses the
 * one its mark wrote at the front of the first paragraph.
 */
export const joinCommentRangesAcrossParagraphs = (
  blocks: readonly BlockContent[],
): number => {
  const startsById = new Map<number, MarkerSlot[]>();
  const endsById = new Map<number, MarkerSlot[]>();

  visitCommentMarkers(blocks, (slot) => {
    const { item } = slot;
    if (item.type === "commentReference") {
      return;
    }
    const slotsById = item.type === "commentRangeStart" ? startsById : endsById;
    const slots = slotsById.get(item.id);
    if (slots) {
      slots.push(slot);
      return;
    }
    slotsById.set(item.id, [slot]);
  });

  const removals = new InlineContentRemovals();
  const blockRemovals = new Map<BlockContent[], number[]>();
  const remove = (slot: MarkerSlot): void => {
    if (slot.level === "inline") {
      removals.mark(slot);
      return;
    }
    const indexes = blockRemovals.get(slot.content) ?? [];
    indexes.push(slot.index);
    blockRemovals.set(slot.content, indexes);
  };
  for (const slots of startsById.values()) {
    for (const slot of slots.slice(1)) {
      remove(slot);
    }
  }
  for (const slots of endsById.values()) {
    for (const slot of slots.slice(0, -1)) {
      remove(slot);
    }
  }
  let dropped = removals.apply();
  for (const [content, indexes] of blockRemovals) {
    for (const index of indexes.toSorted((first, second) => second - first)) {
      content.splice(index, 1);
      dropped += 1;
    }
  }
  return dropped;
};
