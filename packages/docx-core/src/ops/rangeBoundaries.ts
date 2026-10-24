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
 *
 * A paragraph that goes whole takes none of its boundaries with it
 * ({@link paragraphRemains}): bookmarks and move ranges stand between blocks
 * where it was, and comment anchors, which only a paragraph holds, go to the
 * nearest paragraph that stays.
 */

import { panic } from "better-result";

import type {
  BlockContent,
  DocumentBody,
  Paragraph,
  PositionedBookmarkMarker,
} from "../model/document";
import { rejectCopyOriginAncestorsInNodes } from "./copyOriginAncestors";
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

export const rangeBoundaryDefinitions = () => BOUNDARY_SIDES;

type BoundaryType = keyof typeof BOUNDARY_SIDES;
type Boundary = Extract<InlineNode, { type: BoundaryType }>;
type CommentAnchor = Extract<
  Boundary,
  { type: "commentRangeStart" | "commentRangeEnd" }
>;
type OrderedBoundary = Exclude<Boundary, CommentAnchor>;

const isBoundary = (node: InlineNode): node is Boundary =>
  Object.hasOwn(BOUNDARY_SIDES, node.type);

/** The range a boundary belongs to: its family and id. */
const rangeKey = (boundary: Boundary): string =>
  `${BOUNDARY_SIDES[boundary.type].family}:${boundary.id}`;

const isCommentAnchor = (
  node: InlineNode,
): node is Extract<
  Boundary,
  { type: "commentRangeStart" | "commentRangeEnd" }
> => node.type === "commentRangeStart" || node.type === "commentRangeEnd";

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

/** A boundary a block list holds between its blocks. */
export type BlockBoundary = Exclude<Boundary, CommentAnchor>;

// A block list holds no comment anchor, so every boundary in it is one of these.
export const isBlockBoundary = (block: BlockContent): block is BlockBoundary =>
  Object.hasOwn(BOUNDARY_SIDES, block.type);

/** Every range boundary in `content`, at any depth, in document order. */
export const rangeBoundariesIn = (content: readonly InlineNode[]): Boundary[] =>
  content.flatMap((node) =>
    isBoundary(node) ? [node] : rangeBoundariesIn(childNodes(node) ?? []),
  );

type OrderedOccurrence = {
  key: string;
  side: "start" | "end";
  node: OrderedBoundary;
};

const isOrderedBoundary = (
  node: InlineNode | BlockContent,
): node is OrderedBoundary =>
  Object.hasOwn(BOUNDARY_SIDES, node.type) &&
  node.type !== "commentRangeStart" &&
  node.type !== "commentRangeEnd";

const orderedKey = (node: OrderedBoundary): string =>
  `${BOUNDARY_SIDES[node.type].family}:${node.id}`;

const orderedOccurrence = (node: OrderedBoundary): OrderedOccurrence => ({
  key: orderedKey(node),
  side: BOUNDARY_SIDES[node.type].side,
  node,
});

const bookmarkPositions = (
  markers: readonly PositionedBookmarkMarker[] | undefined,
  siblingCount: number,
): ReadonlyMap<number, readonly PositionedBookmarkMarker[]> => {
  const positions = new Map<number, PositionedBookmarkMarker[]>();
  for (const marker of markers ?? []) {
    if (
      !Number.isInteger(marker.index) ||
      marker.index < 0 ||
      marker.index > siblingCount
    )
      continue;
    const atPosition = positions.get(marker.index) ?? [];
    atPosition.push(marker);
    positions.set(marker.index, atPosition);
  }
  return positions;
};

const bookmarkMarkersAt = (
  positions: ReadonlyMap<number, readonly PositionedBookmarkMarker[]>,
  index: number,
): readonly PositionedBookmarkMarker[] => positions.get(index) ?? [];

type InsertsInvertedRangeOptions = {
  body: DocumentBody;
  paragraph: Paragraph;
  candidateContent: readonly InlineNode[];
  incomingContent: readonly InlineNode[];
};

export const insertsInvertedRange = ({
  body,
  paragraph,
  candidateContent,
  incomingContent,
}: InsertsInvertedRangeOptions): boolean => {
  const incoming = new WeakSet<OrderedBoundary>();
  const incomingKeys = new Set<string>();
  const occurrences: OrderedOccurrence[] = [];
  const rememberBoundary = (node: OrderedBoundary): void => {
    incomingKeys.add(orderedKey(node));
    incoming.add(node);
  };
  const rememberIncoming = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isOrderedBoundary(node)) rememberBoundary(node);
      rememberIncoming(childNodes(node) ?? []);
    }
  };
  rememberIncoming(incomingContent);
  if (incomingKeys.size === 0) return false;

  const appendBoundary = (node: OrderedBoundary): void => {
    const occurrence = orderedOccurrence(node);
    if (incomingKeys.has(occurrence.key)) occurrences.push(occurrence);
  };
  const visitInline = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isOrderedBoundary(node)) appendBoundary(node);
      visitInline(childNodes(node) ?? []);
    }
  };
  const visitBlocks = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (isOrderedBoundary(block)) {
        appendBoundary(block);
        continue;
      }
      switch (block.type) {
        case "paragraph":
          visitInline(block === paragraph ? candidateContent : block.content);
          break;
        case "table": {
          const tableBookmarkPositions = bookmarkPositions(
            block.bookmarks,
            block.rows.length,
          );
          for (let rowIndex = 0; rowIndex <= block.rows.length; rowIndex += 1) {
            for (const positioned of bookmarkMarkersAt(
              tableBookmarkPositions,
              rowIndex,
            ))
              appendBoundary(positioned.marker);
            if (rowIndex === block.rows.length) break;
            const row = block.rows[rowIndex];
            if (row === undefined) continue;
            const rowBookmarkPositions = bookmarkPositions(
              row.bookmarks,
              row.cells.length,
            );
            for (
              let cellIndex = 0;
              cellIndex <= row.cells.length;
              cellIndex += 1
            ) {
              for (const positioned of bookmarkMarkersAt(
                rowBookmarkPositions,
                cellIndex,
              ))
                appendBoundary(positioned.marker);
              const cell = row.cells[cellIndex];
              if (cell !== undefined) visitBlocks(cell.content);
            }
          }
          break;
        }
        case "blockSdt":
        case "blockCustomXml":
          visitBlocks(block.content);
          break;
        case "preservedBlock":
          break;
        default: {
          const unreachable: never = block;
          return unreachable;
        }
      }
    }
  };
  visitBlocks(body.content);

  const byKey = new Map<string, OrderedOccurrence[]>();
  for (const occurrence of occurrences) {
    const group = byKey.get(occurrence.key) ?? [];
    group.push(occurrence);
    byKey.set(occurrence.key, group);
  }
  for (const group of byKey.values()) {
    if (group.length !== 2) continue;
    const start = group.findIndex(({ side }) => side === "start");
    const end = group.findIndex(({ side }) => side === "end");
    if (
      start >= 0 &&
      end >= 0 &&
      end < start &&
      group.some(({ node }) => incoming.has(node))
    )
      return true;
  }
  return false;
};

/**
 * What a paragraph that goes whole leaves where it was: every bookmark and
 * move range boundary it holds, at any depth, as blocks in its place, and its
 * comment anchors, in document order, for the nearest paragraph that stays to
 * take. A comment with both anchors there becomes a point there, so the
 * comment and its anchor survive the paragraph.
 */
export const paragraphRemains = (
  content: readonly InlineNode[],
): { all: Boundary[]; blocks: BlockBoundary[]; comments: CommentAnchor[] } => {
  const detached = rejectCopyOriginAncestorsInNodes(content, 0);
  const boundaries = rangeBoundariesIn(detached);
  return {
    all: boundaries,
    blocks: boundaries.filter(
      (boundary): boundary is BlockBoundary => !isCommentAnchor(boundary),
    ),
    comments: boundaries.filter(isCommentAnchor),
  };
};

type ParagraphBoundaryRemovalOptions = {
  blocks: readonly BlockContent[];
  index: number;
  place: "blocks" | "previous";
};

type ParagraphBoundaryRemovalPlan =
  | { type: "blocked" }
  | {
      type: "removed";
      blocks: BlockContent[];
      effects: readonly {
        type: "retained-range-boundaries";
        receiver: Paragraph;
      }[];
    };

export const planParagraphBoundaryRemoval = ({
  blocks,
  index,
  place,
}: ParagraphBoundaryRemovalOptions): ParagraphBoundaryRemovalPlan => {
  const paragraph = blocks[index];
  if (paragraph?.type !== "paragraph")
    return panic("Only a paragraph leaves its range boundaries.");
  const {
    all,
    blocks: boundaries,
    comments,
  } = paragraphRemains(paragraph.content);
  const between = place === "blocks" ? boundaries : [];
  const toPrevious = place === "blocks" ? comments : all;
  const remaining = [
    ...blocks.slice(0, index),
    ...between,
    ...blocks.slice(index + 1),
  ];
  if (toPrevious.length === 0)
    return { type: "removed", blocks: remaining, effects: [] };
  const host = remaining[index - 1];
  if (host?.type !== "paragraph") return { type: "blocked" };
  remaining[index - 1] = { ...host, content: [...host.content, ...toPrevious] };
  return {
    type: "removed",
    blocks: remaining,
    effects: [{ type: "retained-range-boundaries", receiver: host }],
  };
};

/**
 * `nodes` with every deletion `created` names split around the range
 * boundaries directly in it, so they stand outside it, and those
 * boundaries, in document order. Deleting a paragraph takes its content,
 * not a position a user placed in it: accepting the deletion leaves these
 * boundaries where they were, as deleting directly does, and rejecting it
 * gives every piece back in place. A piece takes its deletion's own
 * resolution joins only at the deletion's outer edges.
 */
export const withBoundariesOutsideDeletion = (
  nodes: readonly InlineNode[],
  created: (node: InlineNode) => boolean,
): { nodes: InlineNode[]; standing: Boundary[] } => {
  const standing: Boundary[] = [];
  const visit = (items: readonly InlineNode[]): InlineNode[] =>
    items.flatMap((node) => {
      const children = childNodes(node);
      if (children === undefined) return [node];
      if (!created(node) || node.type !== "deletion") {
        const next = visit(children);
        return next.every((child, index) => child === children[index]) &&
          next.length === children.length
          ? [node]
          : [rebuildNode(node, next)];
      }
      const groups: InlineNode[][] = [[]];
      const between: Boundary[] = [];
      for (const child of children) {
        if (isBoundary(child)) {
          between.push(child);
          groups.push([]);
        } else groups.at(-1)?.push(child);
      }
      if (between.length === 0) return [node];
      standing.push(...between);
      const joins = node.resolutionJoins;
      const last = groups.length - 1;
      const out: InlineNode[] = [];
      for (const [index, group] of groups.entries()) {
        if (group.length > 0) {
          // SAFETY: rebuilding a deletion with other children keeps its type.
          const piece = rebuildNode(node, group) as typeof node;
          if (joins !== undefined)
            piece.resolutionJoins = {
              ...joins,
              before: index === 0 ? joins.before : 0,
              after: index === last ? joins.after : 0,
            };
          out.push(piece);
        }
        const boundary = between[index];
        if (boundary !== undefined) out.push(boundary);
      }
      return out;
    });
  return { nodes: visit(nodes), standing };
};

/** Which kept boundaries each container may hold, as its model declares. */
const HOLDS: Partial<
  Record<InlineNode["type"], (boundary: Boundary) => boolean>
> = {
  inlineSdt: () => true,
  hyperlink: (boundary) =>
    boundary.type === "bookmarkStart" ||
    boundary.type === "bookmarkEnd" ||
    isCommentAnchor(boundary),
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

type RangeMultiplicity =
  | { type: "unpaired" }
  | { type: "ambiguous" }
  | { type: "pair"; start: number; end: number };

export const classifyRangeMultiplicity = (
  occurrences: readonly BoundaryOccurrence[],
  dropped: ReadonlySet<number> = new Set(),
): ReadonlyMap<string, RangeMultiplicity> => {
  const grouped = new Map<string, { starts: number[]; ends: number[] }>();
  for (const [index, occurrence] of occurrences.entries()) {
    if (dropped.has(index)) continue;
    let group = grouped.get(occurrence.range);
    if (group === undefined) {
      group = { starts: [], ends: [] };
      grouped.set(occurrence.range, group);
    }
    const side = occurrence.side === "start" ? group.starts : group.ends;
    side.push(index);
  }
  const classified = new Map<string, RangeMultiplicity>();
  for (const [range, { starts, ends }] of grouped) {
    const start = starts.at(0);
    const end = ends.at(0);
    if (starts.length > 1 || ends.length > 1) {
      classified.set(range, { type: "ambiguous" });
    } else if (start === undefined || end === undefined) {
      classified.set(range, { type: "unpaired" });
    } else {
      classified.set(range, { type: "pair", start, end });
    }
  }
  return classified;
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
    if (!ranges.has(range)) continue;
    const group = copies.get(boundary);
    if (group === undefined) copies.set(boundary, [index]);
    else group.push(index);
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
  const multiplicities = classifyRangeMultiplicity(occurrences, dropped);
  for (const range of ranges) {
    const multiplicity = multiplicities.get(range);
    if (multiplicity?.type !== "pair") continue;
    const { start, end } = multiplicity;
    if (end > start) continue;
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
