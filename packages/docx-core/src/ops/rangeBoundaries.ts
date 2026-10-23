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

import {
  childNodes,
  type InlineNode,
  isEmptyRecord,
  rebuildNode,
} from "./leaves";

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

/**
 * `nodes` without `target`. A run or revision wrapper `target` was all of
 * goes with it, since the model holds no empty one; any other container, a
 * content control or a link, stays, emptied, with its own properties.
 */
const withoutNode = (
  nodes: readonly InlineNode[],
  target: InlineNode,
): readonly InlineNode[] =>
  nodes.flatMap((node) => {
    if (node === target) return [];
    const children = childNodes(node);
    if (
      children === undefined ||
      !children.some((child) => holds(child, target))
    )
      return [node];
    const rebuilt = rebuildNode(node, withoutNode(children, target));
    return isEmptyRecord(rebuilt) ? [] : [rebuilt];
  });

const holds = (node: InlineNode, target: InlineNode): boolean =>
  node === target ||
  (childNodes(node) ?? []).some((child) => holds(child, target));

const withBeside = (
  nodes: readonly InlineNode[],
  anchor: InlineNode,
  inserted: InlineNode,
  side: "before" | "after",
): readonly InlineNode[] =>
  nodes.flatMap((node) => {
    if (node === anchor)
      return side === "after" ? [node, inserted] : [inserted, node];
    const children = childNodes(node);
    if (children === undefined) return [node];
    const next = withBeside(children, anchor, inserted, side);
    return [next === children ? node : rebuildNode(node, next)];
  });

/** One range boundary in document order, as {@link planWellFormedRanges} reads it. */
export type BoundaryOccurrence = {
  /** The marker's element and id: what two copies of one boundary share. */
  boundary: string;
  /** The range it belongs to: its family and id. */
  range: string;
  side: "start" | "end";
  /** Whether a rejection kept it from removed content. */
  kept: boolean;
};

/** What makes kept boundaries' ranges well-formed, by index into the occurrences. */
export type WellFormedRangePlan = {
  dropped: number[];
  /** Each kept boundary that moves to just before or after the other end of its range. */
  moved: { moving: number; anchor: number; side: "before" | "after" }[];
};

/**
 * The one rule that makes every range a rejection kept boundaries of
 * well-formed, for the engine's model and the editor's document alike. A
 * kept copy of a boundary the story still holds is no boundary of its own and
 * goes (of several kept copies, the first stays); then, where such a range's
 * end comes before its start, the half the rejection kept closes on the
 * other: a kept end moves to just after the start, or else a kept start to
 * just before the end. A boundary that was standing never moves, and a range
 * still with more than one start or end is left as it is.
 */
export const planWellFormedRanges = (
  occurrences: readonly BoundaryOccurrence[],
): WellFormedRangePlan => {
  const ranges = new Set(
    occurrences.filter(({ kept }) => kept).map(({ range }) => range),
  );
  const copies = new Map<string, number[]>();
  for (const [index, { boundary, range }] of occurrences.entries()) {
    if (ranges.has(range))
      copies.set(boundary, [...(copies.get(boundary) ?? []), index]);
  }
  const dropped = new Set<number>();
  for (const group of copies.values()) {
    if (group.length < 2) continue;
    const held = group.some((member) => occurrences[member]?.kept === false);
    for (const member of held
      ? group.filter((candidate) => occurrences[candidate]?.kept === true)
      : group.slice(1))
      dropped.add(member);
  }
  const moved: WellFormedRangePlan["moved"] = [];
  for (const range of ranges) {
    const sides = (wanted: "start" | "end") =>
      occurrences.flatMap((occurrence, index) =>
        occurrence.range === range &&
        occurrence.side === wanted &&
        !dropped.has(index)
          ? [index]
          : [],
      );
    const [start, ...moreStarts] = sides("start");
    const [end, ...moreEnds] = sides("end");
    if (start === undefined || end === undefined) continue;
    if (moreStarts.length > 0 || moreEnds.length > 0 || end > start) continue;
    if (occurrences[end]?.kept === true)
      moved.push({ moving: end, anchor: start, side: "after" });
    else if (occurrences[start]?.kept === true)
      moved.push({ moving: start, anchor: end, side: "before" });
  }
  return {
    dropped: [...dropped].toSorted((left, right) => left - right),
    moved,
  };
};

/**
 * The paragraphs' contents with every range a rejection kept boundaries of
 * well-formed, by {@link planWellFormedRanges}. `order` lists the story's
 * paragraph keys in document order. Returns only the paragraphs it changed.
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
  for (const paragraph of order)
    visit(paragraph, contents.get(paragraph) ?? []);
  const plan = planWellFormedRanges(
    found.map(({ node }) => ({
      boundary: `${node.type}:${node.id}`,
      range: rangeKey(node),
      side: BOUNDARY_SIDES[node.type].side,
      kept: relocated.nodes.has(node),
    })),
  );
  const current = (paragraph: string) =>
    changed.get(paragraph) ?? contents.get(paragraph) ?? [];
  for (const index of plan.dropped) {
    const occurrence = found[index];
    if (occurrence !== undefined)
      changed.set(
        occurrence.paragraph,
        withoutNode(current(occurrence.paragraph), occurrence.node),
      );
  }
  for (const { moving: movingIndex, anchor: anchorIndex, side } of plan.moved) {
    const moving = found[movingIndex];
    const anchor = found[anchorIndex];
    if (moving === undefined || anchor === undefined) continue;
    changed.set(
      moving.paragraph,
      withoutNode(current(moving.paragraph), moving.node),
    );
    changed.set(
      anchor.paragraph,
      withBeside(current(anchor.paragraph), anchor.node, moving.node, side),
    );
  }
  return changed;
};
