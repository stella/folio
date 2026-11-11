/**
 * A bookmark owns one range in a story, however many paragraphs it crosses.
 *
 * A bookmark id names one range: `w:bookmarkStart` opens it and the
 * `w:bookmarkEnd` with the same id closes it, wherever in the story they sit.
 * The editor keeps a start whose end it never saw on the paragraph rather
 * than in its text, and writes it back around that paragraph. A paragraph
 * split copies what the paragraph holds to both halves, so the save wrote the
 * range twice under one id, which the format does not allow.
 *
 * This joins them once, for the whole story, as comment ranges are joined:
 * the first start and the last end of an id are the range, and every marker
 * between them is an artefact of the split.
 */

import type { BlockContent, ParagraphContent } from "../types/document";
import { InlineContentRemovals } from "./inlineContentRemovals";
import { visitInlineContentSlots } from "./paragraphTraversal";

type MarkerSlot = { content: ParagraphContent[]; index: number };

const visitBookmarkMarkers = (
  blocks: readonly BlockContent[],
  visit: (slot: MarkerSlot, type: "start" | "end", id: number) => void,
): void => {
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph":
        visitInlineContentSlots(block, ({ content, index, item }) => {
          if (item.type === "bookmarkStart")
            visit({ content, index }, "start", item.id);
          else if (item.type === "bookmarkEnd")
            visit({ content, index }, "end", item.id);
        });
        break;
      case "table":
        for (const row of block.rows)
          for (const cell of row.cells)
            visitBookmarkMarkers(cell.content, visit);
        break;
      case "blockSdt":
      case "blockCustomXml":
        visitBookmarkMarkers(block.content, visit);
        break;
      default:
        break;
    }
  }
};

/**
 * Collapse each bookmark's markers in paragraphs to its first start and its
 * last end. Answers how many markers were dropped.
 */
export const joinBookmarkRangesAcrossParagraphs = (
  blocks: readonly BlockContent[],
): number => {
  const starts = new Map<number, MarkerSlot[]>();
  const ends = new Map<number, MarkerSlot[]>();
  visitBookmarkMarkers(blocks, (slot, type, id) => {
    const byId = type === "start" ? starts : ends;
    const slots = byId.get(id);
    if (slots) slots.push(slot);
    else byId.set(id, [slot]);
  });

  const removals = new InlineContentRemovals();
  for (const slots of starts.values())
    for (const slot of slots.slice(1)) removals.mark(slot);
  for (const slots of ends.values())
    for (const slot of slots.slice(0, -1)) removals.mark(slot);
  return removals.apply();
};
