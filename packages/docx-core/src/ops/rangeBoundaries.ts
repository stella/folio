/**
 * Range boundaries in content a resolution removes.
 *
 * A bookmark, comment or move range boundary is a position, not revision
 * payload. Removing the content that holds one (rejecting an insertion,
 * accepting a deletion, resolving a move) keeps the boundary where that
 * content was, so the range shrinks to what remains and nothing a user placed
 * is dropped. A range the move leaves with its end before its start covers
 * nothing, so its end closes on its start: every range stays well-formed.
 *
 * A comment whose whole range was removed is the comment owner's to retire
 * (see `isCommentExhausted`); its boundaries go with it.
 */

import { childNodes, type InlineNode, rebuildNode } from "./leaves";

const BOUNDARY_SIDES = {
  bookmarkStart: { family: "bookmark", side: "start" },
  bookmarkEnd: { family: "bookmark", side: "end" },
  commentRangeStart: { family: "comment", side: "start" },
  commentRangeEnd: { family: "comment", side: "end" },
  moveFromRangeStart: { family: "moveFrom", side: "start" },
  moveFromRangeEnd: { family: "moveFrom", side: "end" },
  moveToRangeStart: { family: "moveTo", side: "start" },
  moveToRangeEnd: { family: "moveTo", side: "end" },
} as const satisfies Record<string, { family: string; side: "start" | "end" }>;

type BoundaryType = keyof typeof BOUNDARY_SIDES;
type Boundary = Extract<InlineNode, { type: BoundaryType }>;

const isBoundary = (node: InlineNode): node is Boundary =>
  Object.hasOwn(BOUNDARY_SIDES, node.type);

/** The range a boundary belongs to: its family and id. */
const rangeKey = (boundary: Boundary): string =>
  `${BOUNDARY_SIDES[boundary.type].family}:${boundary.id}`;

/**
 * The range boundaries held anywhere in `nodes`, in document order: what
 * stays where removed content was. Each one's range is recorded in
 * `relocated`, for {@link wellFormedRanges}.
 */
export const removedContentBoundaries = (
  nodes: readonly InlineNode[],
  relocated: Set<string>,
): InlineNode[] =>
  nodes.flatMap((node) => {
    if (isBoundary(node)) {
      relocated.add(rangeKey(node));
      return [node];
    }
    return removedContentBoundaries(childNodes(node) ?? [], relocated);
  });

type Occurrence = { paragraph: string; node: Boundary };

const withoutNode = (
  nodes: readonly InlineNode[],
  target: InlineNode,
): readonly InlineNode[] =>
  nodes.flatMap((node) => {
    if (node === target) return [];
    const children = childNodes(node);
    if (children === undefined) return [node];
    const next = withoutNode(children, target);
    return [next === children ? node : rebuildNode(node, next)];
  });

const withAfter = (
  nodes: readonly InlineNode[],
  anchor: InlineNode,
  inserted: InlineNode,
): readonly InlineNode[] =>
  nodes.flatMap((node) => {
    if (node === anchor) return [node, inserted];
    const children = childNodes(node);
    if (children === undefined) return [node];
    const next = withAfter(children, anchor, inserted);
    return [next === children ? node : rebuildNode(node, next)];
  });

/**
 * The paragraphs' contents with every relocated range well-formed: where a
 * range's end now comes before its start, the end moves to just after the
 * start. `order` lists the story's paragraph keys in document order; a range
 * with more than one start or end is left as it is. Returns only the
 * paragraphs it changed.
 */
export const wellFormedRanges = ({
  contents,
  order,
  relocated,
}: {
  contents: ReadonlyMap<string, readonly InlineNode[]>;
  order: readonly string[];
  relocated: ReadonlySet<string>;
}): Map<string, readonly InlineNode[]> => {
  const changed = new Map<string, readonly InlineNode[]>();
  if (relocated.size === 0) return changed;
  const starts = new Map<string, Occurrence[]>();
  const ends = new Map<string, Occurrence[]>();
  const rank = new Map<Boundary, number>();
  const visit = (paragraph: string, nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isBoundary(node)) {
        const key = rangeKey(node);
        if (relocated.has(key)) {
          rank.set(node, rank.size);
          const sides =
            BOUNDARY_SIDES[node.type].side === "start" ? starts : ends;
          sides.set(key, [...(sides.get(key) ?? []), { paragraph, node }]);
        }
        continue;
      }
      visit(paragraph, childNodes(node) ?? []);
    }
  };
  for (const paragraph of order)
    visit(paragraph, contents.get(paragraph) ?? []);
  const current = (paragraph: string) =>
    changed.get(paragraph) ?? contents.get(paragraph) ?? [];
  for (const key of relocated) {
    const [start, ...moreStarts] = starts.get(key) ?? [];
    const [end, ...moreEnds] = ends.get(key) ?? [];
    if (start === undefined || end === undefined) continue;
    if (moreStarts.length > 0 || moreEnds.length > 0) continue;
    if ((rank.get(end.node) ?? 0) > (rank.get(start.node) ?? 0)) continue;
    changed.set(end.paragraph, withoutNode(current(end.paragraph), end.node));
    changed.set(
      start.paragraph,
      withAfter(current(start.paragraph), start.node, end.node),
    );
  }
  return changed;
};
