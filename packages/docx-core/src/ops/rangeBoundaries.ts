/**
 * Range boundaries in content a resolution removes.
 *
 * A bookmark, comment or move range boundary is a position, not revision
 * payload. Rejecting a change removes the content it added (an insertion, a
 * move's destination) but keeps every boundary that content held where it
 * was, so a range shrinks to what remains and nothing a user placed is
 * dropped. A range that leaves ending before its start covers nothing, so
 * its end closes on its start: every range stays well-formed.
 *
 * Accepting a change removes what a direct deletion made earlier would: it
 * keeps the anchors of comment ranges that continue past the removed content
 * and nothing else, so accepting a tracked deletion equals deleting directly.
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

const isCommentAnchor = (node: InlineNode): node is Boundary =>
  node.type === "commentRangeStart" || node.type === "commentRangeEnd";

/** Comment ids with both a start and an end among `nodes`. */
export const pairedCommentRanges = (
  nodes: readonly InlineNode[],
): ReadonlySet<number> => {
  const starts = new Set<number>();
  const ends = new Set<number>();
  const visit = (items: readonly InlineNode[]): void => {
    for (const node of items) {
      if (node.type === "commentRangeStart") starts.add(node.id);
      if (node.type === "commentRangeEnd") ends.add(node.id);
      visit(childNodes(node) ?? []);
    }
  };
  visit(nodes);
  return new Set([...starts].filter((id) => ends.has(id)));
};

/** The boundaries a rejection kept from removed content, and their ranges. */
export type RelocatedBoundaries = {
  ranges: Set<string>;
  nodes: WeakSet<InlineNode>;
};

export const noRelocatedBoundaries = (): RelocatedBoundaries => ({
  ranges: new Set(),
  nodes: new WeakSet(),
});

/**
 * The range boundaries removed content leaves where it was, in document
 * order. Rejecting keeps every one and records it in `relocated`, for
 * {@link wellFormedRanges}; accepting keeps the anchors of
 * `pairedComments`, as a direct deletion does.
 */
export const removedContentBoundaries = ({
  nodes,
  rejecting,
  pairedComments,
  relocated,
}: {
  nodes: readonly InlineNode[];
  rejecting: boolean;
  pairedComments: ReadonlySet<number>;
  relocated: RelocatedBoundaries;
}): InlineNode[] =>
  nodes.flatMap((node) => {
    if (isBoundary(node) && rejecting) {
      relocated.ranges.add(rangeKey(node));
      relocated.nodes.add(node);
      return [node];
    }
    if (isCommentAnchor(node)) return pairedComments.has(node.id) ? [node] : [];
    return removedContentBoundaries({
      nodes: childNodes(node) ?? [],
      rejecting,
      pairedComments,
      relocated,
    });
  });

/** Which kept boundaries each container may hold, as its model declares. */
const HOLDS: Partial<
  Record<InlineNode["type"], (boundary: Boundary) => boolean>
> = {
  inlineSdt: (boundary) => !isCommentAnchor(boundary),
  hyperlink: (boundary) =>
    boundary.type === "bookmarkStart" || boundary.type === "bookmarkEnd",
  inlineWrapper: () => true,
  insertion: () => true,
  deletion: () => true,
  moveFrom: () => true,
  moveTo: () => true,
};

/**
 * `node` with `boundaries` at its start or end, as deep as `depth` containers
 * in as the containers on the way can hold them: where removed content that
 * cut those containers was, so the pieces it separated still join around the
 * boundaries, down to the level that holds them. `undefined` when `node`
 * itself cannot hold them, or `depth` is 0.
 */
export const withBoundariesAtEdge = (
  node: InlineNode,
  boundaries: readonly InlineNode[],
  depth: number,
  edge: "start" | "end",
): InlineNode | undefined => {
  if (depth <= 0 || boundaries.length === 0) return undefined;
  const children = childNodes(node);
  const holds = HOLDS[node.type];
  if (
    children === undefined ||
    holds === undefined ||
    !boundaries.every((boundary) => isBoundary(boundary) && holds(boundary))
  )
    return undefined;
  const inner = edge === "start" ? children.at(0) : children.at(-1);
  const nested =
    depth > 1 && inner !== undefined
      ? withBoundariesAtEdge(inner, boundaries, depth - 1, edge)
      : undefined;
  if (nested !== undefined)
    return rebuildNode(
      node,
      edge === "start"
        ? [nested, ...children.slice(1)]
        : [...children.slice(0, -1), nested],
    );
  return rebuildNode(
    node,
    edge === "start"
      ? [...boundaries, ...children]
      : [...children, ...boundaries],
  );
};

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
 * The paragraphs' contents with every range a rejection kept boundaries of
 * well-formed. A kept copy of a boundary the story still holds is no boundary
 * of its own and goes (of several kept copies, the first stays); then, where
 * a range's end comes before its start, the end moves to just after the
 * start. `order` lists the story's paragraph keys in document order; a range
 * still with more than one start or end is left as it is. Returns only the
 * paragraphs it changed.
 */
export const wellFormedRanges = ({
  contents,
  order,
  relocated,
}: {
  contents: ReadonlyMap<string, readonly InlineNode[]>;
  order: readonly string[];
  relocated: RelocatedBoundaries;
}): Map<string, readonly InlineNode[]> => {
  const changed = new Map<string, readonly InlineNode[]>();
  if (relocated.ranges.size === 0) return changed;
  const current = (paragraph: string) =>
    changed.get(paragraph) ?? contents.get(paragraph) ?? [];
  const occurrences = (): Occurrence[] => {
    const found: Occurrence[] = [];
    const visit = (paragraph: string, nodes: readonly InlineNode[]): void => {
      for (const node of nodes) {
        if (isBoundary(node)) {
          if (relocated.ranges.has(rangeKey(node)))
            found.push({ paragraph, node });
          continue;
        }
        visit(paragraph, childNodes(node) ?? []);
      }
    };
    for (const paragraph of order) visit(paragraph, current(paragraph));
    return found;
  };
  const copies = new Map<string, Occurrence[]>();
  for (const occurrence of occurrences()) {
    const key = `${occurrence.node.type}:${occurrence.node.id}`;
    copies.set(key, [...(copies.get(key) ?? []), occurrence]);
  }
  for (const group of copies.values()) {
    if (group.length < 2) continue;
    const held = group.some(({ node }) => !relocated.nodes.has(node));
    const dropped = held
      ? group.filter(({ node }) => relocated.nodes.has(node))
      : group.slice(1);
    for (const { paragraph, node } of dropped)
      changed.set(paragraph, withoutNode(current(paragraph), node));
  }
  const remaining = occurrences();
  const side = (key: string, wanted: "start" | "end") =>
    remaining.flatMap((occurrence, rank) =>
      rangeKey(occurrence.node) === key &&
      BOUNDARY_SIDES[occurrence.node.type].side === wanted
        ? [{ paragraph: occurrence.paragraph, node: occurrence.node, rank }]
        : [],
    );
  for (const key of relocated.ranges) {
    const starts = side(key, "start");
    const ends = side(key, "end");
    const [start] = starts;
    const [end] = ends;
    if (start === undefined || end === undefined) continue;
    if (starts.length > 1 || ends.length > 1 || end.rank > start.rank) continue;
    changed.set(end.paragraph, withoutNode(current(end.paragraph), end.node));
    changed.set(
      start.paragraph,
      withAfter(current(start.paragraph), start.node, end.node),
    );
  }
  return changed;
};
