/**
 * A story projection's addresses held per top-level block, so the next
 * projection can keep every block an edit left and walk only the ones it
 * changed.
 *
 * A paragraph's start is kept relative to the top-level node holding it, and
 * the paragraph ids are indexed by block, so the blocks after an edit keep
 * their records: their starts follow from the new block starts and their
 * index entries from the number of blocks the edit added or removed.
 */
import type { Node as PMNode } from "prosemirror-model";

import {
  type ParagraphCommentMarkers,
  visitCommentMarkerItems,
} from "../docx/commentAnchorIndex";
import { visitBookmarkBoundaries } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, Paragraph } from "../types/document";
import type {
  CanonicalInlineGap,
  CanonicalSeamExtent,
} from "./canonicalInlineProjection";

export type ParagraphAddress = {
  blockId: string;
  container: readonly BlockContent[];
  start: number;
  text: string;
  boundaries: readonly (readonly CanonicalInlineGap[])[];
  seams: readonly (readonly CanonicalSeamExtent[])[];
  node: PMNode;
  source: Paragraph;
};

/** A paragraph as the top-level block holding it projects it. */
export type BlockParagraph = {
  blockId: string;
  /** The blocks holding the paragraph, or undefined for the story's own. */
  container: readonly BlockContent[] | undefined;
  /** From the start of the top-level node holding it. */
  offset: number;
  inline: Pick<ParagraphAddress, "text" | "boundaries" | "seams">;
  node: PMNode;
  source: Paragraph;
};

/** What projecting one top-level node found. */
export type ProjectedBlock = {
  /** The node the projection walked. */
  node: PMNode;
  /** The story block at the node's index, which the node was projected from when the projection is reusable. */
  source: BlockContent | undefined;
  paragraphs: readonly BlockParagraph[];
  /** The id of every source paragraph the block holds, addressed or not. */
  sourceIds: readonly string[];
  /** The comment ranges open at the block's start and after it, in their set's order. */
  openBefore: readonly number[];
  openAfter: readonly number[];
};

/** The story-wide facts every paragraph's gaps read. */
export type StoryProjectionFacts = {
  pairedBookmarkIds: ReadonlySet<number>;
  rangedIds: ReadonlySet<number>;
  pointIds: ReadonlySet<number>;
};

/** Where a paragraph id stands: its block, and its paragraph there or -1 when it has no address. */
export type ParagraphSlot = { block: number; paragraph: number };

/** Index levels a lookup passes before the index is rebuilt flat. */
const MAX_INDEX_DEPTH = 8;

/**
 * Paragraph ids by the block holding them. An index for the next projection
 * keeps this one as its base and records only the blocks the edit replaced,
 * shifting the slots of the blocks after them, so building it costs the
 * replaced blocks and a lookup at most a few levels.
 */
type IndexBase = {
  index: ParagraphIdIndex;
  /** Ids of the base's replaced blocks. */
  removed: ReadonlySet<string>;
  /** The base's blocks from here on stand `shift` blocks later. */
  end: number;
  shift: number;
};

export class ParagraphIdIndex {
  private readonly own: ReadonlyMap<string, ParagraphSlot>;
  private readonly base: IndexBase | undefined;
  private readonly depth: number;

  private constructor(
    own: ReadonlyMap<string, ParagraphSlot>,
    base: IndexBase | undefined,
    depth: number,
  ) {
    this.own = own;
    this.base = base;
    this.depth = depth;
  }

  /** Every id of `blocks`; an addressed paragraph's slot over an unaddressed one of the same id. */
  static of(blocks: readonly ProjectedBlock[]): ParagraphIdIndex {
    return new ParagraphIdIndex(slotsOf(blocks, 0), undefined, 0);
  }

  get(id: string): ParagraphSlot | undefined {
    const own = this.own.get(id);
    if (own !== undefined) return own;
    const base = this.base;
    if (base === undefined || base.removed.has(id)) return undefined;
    const slot = base.index.get(id);
    if (slot === undefined || slot.block < base.end) return slot;
    return { block: slot.block + base.shift, paragraph: slot.paragraph };
  }

  /**
   * This index with its blocks from `start` up to `end` replaced by
   * `window`, the blocks after them moved by the difference. `blocks` is the
   * new projection's every block, read when the index is rebuilt flat.
   */
  replaced({
    start,
    end,
    removed,
    window,
    blocks,
  }: {
    start: number;
    end: number;
    removed: ReadonlySet<string>;
    window: readonly ProjectedBlock[];
    blocks: readonly ProjectedBlock[];
  }): ParagraphIdIndex {
    if (this.depth + 1 >= MAX_INDEX_DEPTH) return ParagraphIdIndex.of(blocks);
    return new ParagraphIdIndex(
      slotsOf(window, start),
      { index: this, removed, end, shift: start + window.length - end },
      this.depth + 1,
    );
  }
}

const slotsOf = (
  blocks: readonly ProjectedBlock[],
  first: number,
): Map<string, ParagraphSlot> => {
  const slots = new Map<string, ParagraphSlot>();
  for (const [index, block] of blocks.entries())
    for (const id of block.sourceIds)
      slots.set(id, { block: first + index, paragraph: -1 });
  for (const [index, block] of blocks.entries())
    for (const [paragraph, { blockId }] of block.paragraphs.entries())
      slots.set(blockId, { block: first + index, paragraph });
  return slots;
};

/** Whether two id lists hold the same ids in the same order. */
export const sameIdOrder = (
  left: readonly number[],
  right: readonly number[],
): boolean =>
  left.length === right.length && left.every((id, at) => right[at] === id);

/**
 * The story-wide facts `blocks` contribute, in document order: each bookmark
 * boundary, each comment range start and each comment reference. Two runs of
 * blocks contributing the same sequence leave the story's facts as they were
 * when one stands in for the other.
 */
export const storyFactContributions = (
  blocks: readonly BlockContent[],
  markersOf: ParagraphCommentMarkers,
): number[] => {
  const contributions: number[] = [];
  visitBookmarkBoundaries(blocks, (id, type) =>
    contributions.push(type === "start" ? 0 : 1, id),
  );
  visitCommentMarkerItems(blocks, markersOf, (item) => {
    if (item.type === "commentRangeStart") contributions.push(2, item.id);
    if (item.type === "commentReference") contributions.push(3, item.id);
  });
  return contributions;
};
