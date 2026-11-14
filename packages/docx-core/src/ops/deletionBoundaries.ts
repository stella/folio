/**
 * Which range boundaries a direct deletion keeps in place.
 *
 * A comment range and a tracked move range are each a start and an end; one
 * without the other is invalid, and the save refuses a document holding one.
 * A deletion whose range holds one end of such a range and not the other
 * keeps that end where it stands, so the range shrinks to the deletion's
 * edge. A deletion that holds both ends takes the whole range: a comment
 * whose anchored text is gone is then retired with its reference by the
 * comment exhaustion rule (`survivingComments`). A range that is already
 * unbalanced or inverted in the story is not this deletion's to repair, and its
 * boundaries go as the rest of the deleted content does, so an insertion of
 * one is still undone by deleting it.
 */

import type { BlockContent, DocumentBody, Paragraph } from "../model/document";
import { compareGaps, type Gap, type LeafSpan, leafSpans } from "./leaves";
import {
  isBlockBoundary,
  isBoundary,
  rangeBoundariesIn,
  rangeBoundaryDefinitions,
} from "./rangeBoundaries";

/** The families whose ranges must stay balanced for a save to be valid. */
const BALANCED_FAMILIES: ReadonlySet<string> = new Set([
  "comment",
  "moveFrom",
  "moveTo",
]);

type Occurrence = { node: object; side: "start" | "end" };

const BOUNDARY_SIDES = rangeBoundaryDefinitions();

/** `family:id` and side of a boundary of a family that must stay balanced. */
const balancedRangeOf = (node: {
  type: string;
  id?: unknown;
}): { key: string; side: "start" | "end" } | undefined => {
  if (!Object.hasOwn(BOUNDARY_SIDES, node.type)) return undefined;
  // SAFETY: `Object.hasOwn` above proved the type is one of the table's keys.
  const { family, side } =
    BOUNDARY_SIDES[node.type as keyof typeof BOUNDARY_SIDES];
  if (!BALANCED_FAMILIES.has(family)) return undefined;
  return { key: `${family}:${String(node.id)}`, side };
};

const record = (
  ranges: Map<string, Occurrence[]>,
  node: { type: string; id?: unknown },
): void => {
  const range = balancedRangeOf(node);
  if (range === undefined) return;
  const occurrences = ranges.get(range.key) ?? [];
  occurrences.push({ node, side: range.side });
  ranges.set(range.key, occurrences);
};

/** Whether each start is followed by its own end, in document order. */
const wellFormed = (occurrences: readonly Occurrence[]): boolean => {
  let open = 0;
  for (const { side } of occurrences) {
    open += side === "start" ? 1 : -1;
    if (open < 0) return false;
  }
  return open === 0;
};

const recordBlocks = (
  blocks: readonly BlockContent[],
  ranges: Map<string, Occurrence[]>,
): void => {
  for (const block of blocks) {
    if (isBlockBoundary(block)) {
      record(ranges, block);
      continue;
    }
    switch (block.type) {
      case "paragraph":
        for (const boundary of rangeBoundariesIn(block.content))
          record(ranges, boundary);
        break;
      case "table":
        for (const row of block.rows)
          for (const cell of row.cells) recordBlocks(cell.content, ranges);
        break;
      case "blockSdt":
      case "blockCustomXml":
        recordBlocks(block.content, ranges);
        break;
      case "preservedBlock":
        break;
      default: {
        const unhandled: never = block;
        return unhandled;
      }
    }
  }
};

/**
 * The boundaries of `paragraph` between `from` and `to` that a direct
 * deletion keeps: each one whose range is balanced in `body` and would not
 * be without the boundaries the deletion covers. A range is well formed
 * when each start is followed by its own end. In document order.
 */
export const keptDeletionBoundaries = ({
  body,
  paragraph,
  from,
  to,
}: {
  body: Pick<DocumentBody, "content">;
  paragraph: Paragraph;
  from: Gap;
  to: Gap;
}): LeafSpan[] => {
  const covered = leafSpans(paragraph.content).filter(
    ({ node, before, after }) =>
      isBoundary(node) &&
      balancedRangeOf(node) !== undefined &&
      compareGaps(from, before) <= 0 &&
      compareGaps(after, to) <= 0,
  );
  if (covered.length === 0) return [];
  const removed = new Set<object>(covered.map(({ node }) => node));
  const ranges = new Map<string, Occurrence[]>();
  recordBlocks(body.content, ranges);
  const unbalancedByDeletion = new Set(
    [...ranges].flatMap(([key, occurrences]) =>
      occurrences.some(({ node }) => removed.has(node)) &&
      wellFormed(occurrences) &&
      !wellFormed(occurrences.filter(({ node }) => !removed.has(node)))
        ? [key]
        : [],
    ),
  );
  return covered.filter(({ node }) => {
    const range = balancedRangeOf(node);
    return range !== undefined && unbalancedByDeletion.has(range.key);
  });
};
