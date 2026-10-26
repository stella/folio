/**
 * What survives when an edit replaces a span of inline content.
 *
 * ProseMirror gives a replacement the marks it inherits at the boundary, and
 * drops any mark declared `inclusive: false` once the span reaches the end of
 * what that mark covers. For formatting that is right. For a mark that names
 * something outside itself — the comment this text is commented by, the
 * address this text links to — it is not: the target survives the edit, and
 * only the mark saying so is lost. A comment loses the range start that lived
 * on the replaced text, so the save emits a `commentRangeEnd` with no
 * `commentRangeStart`, invalid OOXML and a comment anchored to nothing; a
 * hyperlink simply stops being a link. Zero-width anchors fare worse still: a
 * plain replace deletes them outright.
 *
 * Text replacement preserves a comment's aligned coverage. Unchanged text
 * keeps its ranges; inserted text takes the comments of the replaced aligned
 * span. A range with no remaining content becomes a point at the replacement
 * start. The applier decides alignment once for direct and tracked edits.
 *
 * - `comment` mark: carried over its aligned replacement span.
 * - `hyperlink` mark: the replacement carries the first link the span held, so
 *   replacing a link's display text leaves the link on the new text. A span
 *   crossing two links keeps the one it opens in, because the mark excludes
 *   itself and a replacement is one run of text with one address.
 * - `footnoteRef` mark (footnotes and endnotes alike): not carried. It marks
 *   the reference's own number, not prose around it, and it serializes as a
 *   bare `w:footnoteReference` run, so carrying it onto a replacement would
 *   discard the new text. Replacing a reference's text deletes the reference,
 *   as deleting that text does.
 * - `runIdentity` mark: not carried. It is one authored run's identity and the
 *   markup that run carried, not a target. A replacement is characters the
 *   session named by its `w:rsidR` did not write, so carrying it would fuse
 *   the replacement into a run it was never part of and stamp it with an
 *   editing session it did not belong to.
 * - tracked-change marks (`insertion`, `deletion`, `runPropertyChange`): not
 *   carried. The applier decides them per edit mode and writes them itself.
 * - `inlineWrapper` mark: not carried. Every wrapper kind states something
 *   about the text under it and names nothing outside itself, so boundary
 *   inheritance is the right rule: a replacement inside a bidirectional
 *   override is still inside it, and one that replaces the whole of it is not.
 * - `commentReference` node: kept after the replacement, where the comment's
 *   last covered position now is.
 * - `rangeAnchor` node: kept after the replacement. A comment or move range
 *   that spans no content has no mark to carry, so replacing the words around
 *   a point comment would otherwise delete the comment outright.
 * - `moveRangeBoundary` node: a start kept before the replacement, an end
 *   after it, so replacing a moved span keeps that range's identity and scope.
 * - `bookmarkBoundary` node: a start kept before the replacement, an end after
 *   it, so the bookmark still covers the new text.
 * - `textBoxAnchor` node: kept after the replacement; the shape it anchors
 *   outlives the sentence it was typed into.
 * - `renderedPageBreak` node: dropped with the text. It is Word's cached
 *   pagination, re-derived on layout, and nothing else names it.
 * - every other mark: formatting, and left to boundary inheritance, which
 *   already gives the replacement the formatting of the text it replaces.
 *
 * Removing the owning block removes its comment. Replacing its last covered
 * text within a surviving block retains a point rather than dropping it.
 */

import {
  Fragment,
  Mark,
  type Node as PMNode,
  type Schema,
} from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { isCommentExhausted } from "@stll/docx-core/ops";

import { isCommentAnchorCarrier } from "../ai-edits/comment-lifecycle";
import {
  expectParagraphAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
} from "./attrs";
import { expectCommentRangeBoundaryAttrs } from "./commentRangeBoundaryAttrs";
import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { COMMENT_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/CommentRangeBoundaryExtension";
import { COMMENT_REFERENCE_NODE_NAME } from "./extensions/nodes/CommentReferenceExtension";
import { MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/MoveRangeBoundaryExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import { TEXT_BOX_ANCHOR_NODE_NAME } from "./extensions/nodes/TextBoxAnchorExtension";
import { readMoveRangeBoundaryAttrs } from "./moveRangeBoundaryAttrs";
import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";
import { isZeroWidthAnchor } from "./zeroWidthAnchors";

/**
 * What a replacement does with each mark the schema declares
 * `inclusive: false`. Boundary inheritance already handles every other mark,
 * so this table is exactly the set that needs a decision; a new
 * non-inclusive mark that is missing from it fails
 * `replacedAnnotations.test.ts` rather than silently following the text.
 */
export const NON_INCLUSIVE_MARK_DISPOSITION: {
  readonly comment: "carry";
  readonly hyperlink: "carry";
  readonly footnoteRef: "followsTheText";
  readonly runIdentity: "followsTheText";
  readonly insertion: "followsTheText";
  readonly deletion: "followsTheText";
  readonly runPropertyChange: "followsTheText";
  readonly inlineWrapper: "followsTheText";
} = {
  comment: "carry",
  hyperlink: "carry",
  footnoteRef: "followsTheText",
  runIdentity: "followsTheText",
  insertion: "followsTheText",
  deletion: "followsTheText",
  runPropertyChange: "followsTheText",
  inlineWrapper: "followsTheText",
} as const satisfies Record<string, "carry" | "followsTheText">;

/** The mark names in the schema that a replacement must not silently drop. */
export const nonInclusiveMarkNames = (schema: Schema): string[] =>
  Object.entries(schema.marks)
    .filter(([, type]) => type.spec.inclusive === false)
    .map(([name]) => name);

const isCarried = (mark: Mark): boolean =>
  NON_INCLUSIVE_MARK_DISPOSITION[
    mark.type.name as keyof typeof NON_INCLUSIVE_MARK_DISPOSITION
  ] === "carry";

/** One typed anchor for a comment authored as a point. */
export const commentPoint = (schema: Schema, id: number): PMNode =>
  schema.node(RANGE_ANCHOR_NODE_NAME, {
    start: { type: "commentRangeStart", id },
    end: { type: "commentRangeEnd", id },
  });

type ResolvedCommentPointsOptions = {
  doc: PMNode;
  removes: (node: PMNode, position: number) => boolean;
  retiresStructure: (node: PMNode, position: number) => boolean;
};

/** The same collector serves standalone commands and the existing bulk walk. */
export const createResolvedCommentPointCollector = ({
  removes,
  retiresStructure,
}: Omit<ResolvedCommentPointsOptions, "doc">): {
  visit: (node: PMNode, position: number, previous?: PMNode) => boolean;
  finish: () => Map<number, number>;
} => {
  const points = new Map<number, number>();
  const retained = new Set<number>();
  const removedStructures: { from: number; to: number }[] = [];
  const visit = (
    node: PMNode,
    position: number,
    previous?: PMNode,
  ): boolean => {
    if (node.type.name === "paragraph") {
      const mark = expectParagraphAttrs(node).pPrMark;
      // A final block deletion retires the boundary on its preceding carrier.
      const precedingMark =
        previous?.type.name === "paragraph"
          ? expectParagraphAttrs(previous).pPrMark
          : null;
      const retiredBoundary =
        ((mark?.kind === "del" || mark?.kind === "moveFrom") &&
          retiresStructure(node, position)) ||
        ((precedingMark?.kind === "del" ||
          precedingMark?.kind === "moveFrom") &&
          previous !== undefined &&
          retiresStructure(previous, position - previous.nodeSize));
      if (retiredBoundary) {
        let survivingContent = false;
        node.descendants((inline, offset) => {
          if (
            inline.isInline &&
            !isCommentAnchorCarrier(inline) &&
            !removes(inline, position + 1 + offset)
          )
            survivingContent = true;
          return !survivingContent;
        });
        // Keep walking an emptied carrier so its ranged comments are recorded
        // for removal along with the carrier.
        if (!survivingContent)
          removedStructures.push({
            from: position,
            to: position + node.nodeSize,
          });
      }
    }
    if (
      node.type.name === "tableRow" &&
      expectTableRowAttrs(node).trDel &&
      retiresStructure(node, position)
    ) {
      removedStructures.push({ from: position, to: position + node.nodeSize });
      return true;
    }
    if (
      (node.type.name === "tableCell" || node.type.name === "tableHeader") &&
      expectTableCellAttrs(node).cellMarker?.kind === "del" &&
      retiresStructure(node, position)
    ) {
      removedStructures.push({ from: position, to: position + node.nodeSize });
      return true;
    }
    if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
      const attrs = readRangeAnchorAttrs(node);
      if (attrs.ok && attrs.value.start.type === "commentRangeStart")
        retained.add(attrs.value.start.id);
    }
    if (!node.isInline || isCommentAnchorCarrier(node)) return true;
    const removedWithStructure = removedStructures.some(
      ({ from, to }) => position > from && position < to,
    );
    for (const mark of node.marks) {
      if (mark.type.name !== "comment") continue;
      const id = mark.attrs["commentId"];
      if (!removedWithStructure && !removes(node, position)) retained.add(id);
      else if (!points.has(id)) points.set(id, position);
    }
    return true;
  };
  const finish = (): Map<number, number> => {
    for (const id of points.keys()) {
      if (
        !isCommentExhausted({
          type: "coverage",
          anchorKind: "range",
          survivingNonCarrierContent: retained.has(id),
        })
      )
        points.delete(id);
    }
    return points;
  };
  return { visit, finish };
};

/** Ranged comments whose covered non-carrier content is removed by a resolution. */
export const resolvedCommentPoints = ({
  doc,
  removes,
  retiresStructure,
}: ResolvedCommentPointsOptions): Map<number, number> => {
  const collector = createResolvedCommentPointCollector({
    removes,
    retiresStructure,
  });
  doc.descendants((node, position, parent, index) =>
    collector.visit(
      node,
      position,
      index > 0 ? (parent?.maybeChild(index - 1) ?? undefined) : undefined,
    ),
  );
  return collector.finish();
};

/** Anchors an edit must put back, and where relative to the replacement. */
const ANCHOR_PLACEMENT = {
  [COMMENT_REFERENCE_NODE_NAME]: "after",
  [COMMENT_RANGE_BOUNDARY_NODE_NAME]: "byCommentRangeBoundaryType",
  [RANGE_ANCHOR_NODE_NAME]: "after",
  [TEXT_BOX_ANCHOR_NODE_NAME]: "after",
  [BOOKMARK_BOUNDARY_NODE_NAME]: "byBoundaryType",
  [MOVE_RANGE_BOUNDARY_NODE_NAME]: "byMoveBoundaryType",
} as const;

type AnchorPlacement = (typeof ANCHOR_PLACEMENT)[keyof typeof ANCHOR_PLACEMENT];

const anchorPlacement = (node: PMNode): AnchorPlacement | undefined =>
  ANCHOR_PLACEMENT[node.type.name as keyof typeof ANCHOR_PLACEMENT];

const opensRange = (node: PMNode, placement: AnchorPlacement): boolean => {
  if (placement === "byBoundaryType") {
    return node.attrs["type"] === "start";
  }
  if (placement === "byCommentRangeBoundaryType") {
    return (
      expectCommentRangeBoundaryAttrs(node).marker.type === "commentRangeStart"
    );
  }
  if (placement !== "byMoveBoundaryType") {
    return false;
  }
  const marker = readMoveRangeBoundaryAttrs(node);
  return marker.ok && marker.value.type.endsWith("Start");
};

export type ReplacedAnnotations = {
  /** Marks the replacement must carry, in document order. */
  carried: readonly Mark[];
  /** Anchors restored before the replacement, in their original order. */
  leading: readonly PMNode[];
  /** Anchors restored after the replacement, in their original order. */
  trailing: readonly PMNode[];
};

const NOTHING_REPLACED: ReplacedAnnotations = {
  carried: [],
  leading: [],
  trailing: [],
};

export const hasReplacedAnnotations = ({
  carried,
  leading,
  trailing,
}: ReplacedAnnotations): boolean =>
  carried.length > 0 || leading.length > 0 || trailing.length > 0;

/**
 * The annotations carried by `[from, to)`. An anchor counts only when the
 * range covers it whole: a partially covered atom is not deleted by the
 * replacement either.
 */
export const surveyReplacedAnnotations = (
  doc: PMNode,
  from: number,
  to: number,
): ReplacedAnnotations => {
  if (to <= from) {
    return NOTHING_REPLACED;
  }

  const carried: Mark[] = [];
  const leading: PMNode[] = [];
  const trailing: PMNode[] = [];

  doc.nodesBetween(from, to, (node, position) => {
    if (!node.isInline) {
      return true;
    }
    for (const mark of node.marks) {
      if (isCarried(mark) && !alreadyCarried(mark, carried)) {
        carried.push(mark);
      }
    }
    const placement = anchorPlacement(node);
    if (
      placement === undefined ||
      position < from ||
      position + node.nodeSize > to
    ) {
      return false;
    }
    (opensRange(node, placement) ? leading : trailing).push(node);
    return false;
  });

  return { carried, leading, trailing };
};

/**
 * Whether the span already contributed this mark, or one the mark would evict.
 * Comments do not exclude each other, so every comment is collected; a
 * hyperlink excludes its own type, so the first address in the span wins
 * rather than the last silently overwriting it.
 */
const alreadyCarried = (mark: Mark, carried: readonly Mark[]): boolean =>
  carried.some((other) => other.eq(mark) || mark.type.excludes(other.type));

type AnnotatedReplacementOptions = {
  annotations: ReplacedAnnotations;
  /** The replacement's own inline content, already carrying its formatting. */
  content: Fragment;
};

type AnnotatedReplacement = {
  fragment: Fragment;
  /** Size of the anchors restored before the replacement's own content. */
  leadingSize: number;
  /** Size of the replacement's own content, once the marks are on it. */
  contentSize: number;
};

const anchorsSize = (anchors: readonly PMNode[]): number => {
  let size = 0;
  for (const node of anchors) {
    size += node.nodeSize;
  }
  return size;
};

/** `marks` with the carried marks added, each type's exclusions respected. */
export const addCarriedMarks = (
  marks: readonly Mark[],
  carried: readonly Mark[],
): readonly Mark[] => {
  let next = marks;
  for (const mark of carried) {
    next = mark.addToSet(next);
  }
  return next;
};

/**
 * Invariant: a comment mark covers ONE contiguous stretch of a paragraph. A
 * save writes one range start and one range end per comment, so a comment on
 * two stretches reopens covering the text between them as well.
 *
 * Text an edit inserts is where that breaks: it lands next to the text of a
 * comment, or inside it. These pick the comments inserted text takes so that
 * every comment stays one stretch.
 */

/**
 * The first (`1`) or last (`-1`) inline content inside `node` that is not a
 * zero-width anchor, descending through every container.
 */
const innermostEdge = (node: PMNode, direction: -1 | 1): PMNode | null => {
  if (node.isInline) {
    return isZeroWidthAnchor(node) ? null : node;
  }
  for (
    let index = direction === 1 ? 0 : node.childCount - 1;
    index >= 0 && index < node.childCount;
    index += direction
  ) {
    const found = innermostEdge(node.child(index), direction);
    if (found) {
      return found;
    }
  }
  return null;
};

/**
 * The inline content nearest the block boundary `at` on one side, in document
 * order: past empty paragraphs, and into and out of tables, the way a saved
 * range's two points read it.
 */
const contentBesideBoundary = (
  doc: PMNode,
  at: number,
  direction: -1 | 1,
): PMNode | null => {
  const $at = doc.resolve(at);
  for (let depth = $at.depth; depth >= 0; depth--) {
    const parent = $at.node(depth);
    const index = $at.index(depth);
    // At the boundary's own depth `index` is the child after it; above it, the
    // child that contains the boundary, which lies on neither side.
    let child = index + direction;
    if (depth === $at.depth) {
      child = direction === 1 ? index : index - 1;
    }
    for (; child >= 0 && child < parent.childCount; child += direction) {
      const found = innermostEdge(parent.child(child), direction);
      if (found) {
        return found;
      }
    }
  }
  return null;
};

/**
 * The content either side of `at` in its paragraph, past zero-width anchors
 * (a comment's own reference sits between its text and what follows).
 */
const neighboursAt = (
  doc: PMNode,
  at: number,
): { before: PMNode | null; after: PMNode | null } => {
  const $at = doc.resolve(at);
  const parent = $at.parent;
  const neighbour = (direction: -1 | 1): PMNode | null => {
    const adjacent = direction === -1 ? $at.nodeBefore : $at.nodeAfter;
    if (!adjacent || !isZeroWidthAnchor(adjacent)) {
      return adjacent;
    }
    // An anchor is an atom, so `at` sits on a child boundary here.
    for (
      let index = direction === -1 ? $at.index() - 2 : $at.index() + 1;
      index >= 0 && index < parent.childCount;
      index += direction
    ) {
      const child = parent.child(index);
      if (!isZeroWidthAnchor(child)) {
        return child;
      }
    }
    return null;
  };
  return { before: neighbour(-1), after: neighbour(1) };
};

/** The comments on both `before` and `after`: text between them must carry them. */
const commentsAround = (before: PMNode | null, after: PMNode | null): Mark[] =>
  before && after
    ? before.marks.filter(
        (mark) => mark.type.name === "comment" && mark.isInSet(after.marks),
      )
    : [];

/**
 * The mark sets of a run of text a TRACKED edit inserts at `at` (one set per
 * piece, in order), with the comments that keep every comment one stretch.
 *
 * A tracked replacement keeps the replaced text (as a deletion), and that text
 * keeps its comments, so a comment carried or inherited onto the insertion
 * only CONTINUES its range: it stays when, through the run, it reaches the
 * content directly before or after `at`, and goes otherwise. Between two
 * pieces that have a comment (the neighbours included), every piece takes
 * it, or the run would split it in two.
 */
export const commentsForTrackedInsertion = (
  doc: PMNode,
  at: number,
  pieces: readonly (readonly Mark[])[],
): (readonly Mark[])[] => {
  const { before, after } = neighboursAt(doc, at);
  const sets: (readonly Mark[])[] = [
    before?.marks ?? [],
    ...pieces,
    after?.marks ?? [],
  ];
  const comments: Mark[] = [];
  for (const set of sets) {
    for (const mark of set) {
      if (mark.type.name === "comment" && !mark.isInSet(comments)) {
        comments.push(mark);
      }
    }
  }
  const result: (readonly Mark[])[] = pieces.map((set) =>
    set.filter((mark) => mark.type.name !== "comment"),
  );
  for (const comment of comments) {
    const covered = sets.flatMap((set, index) =>
      comment.isInSet(set) ? [index] : [],
    );
    const first = covered[0]!;
    const last = covered.at(-1)!;
    // Only a stretch that reaches a neighbour continues the comment.
    if (first !== 0 && last !== sets.length - 1) {
      continue;
    }
    for (
      let index = Math.max(first, 1);
      index <= Math.min(last, pieces.length);
      index++
    ) {
      result[index - 1] = comment.addToSet([...result[index - 1]!]);
    }
  }
  return result;
};

/**
 * Where, within the text a tracked replacement removes (`from` to `to`), each
 * piece of its new text goes, one position per piece and never decreasing.
 *
 * The redline reads the removed text first and the new text after it, so by
 * default every piece goes at `to`. A piece given a comment goes where that
 * comment's stretch of removed text ends instead, next to it; the pieces
 * before it that have no comment go where that stretch begins, and the ones
 * after it follow it. So each comment stays one stretch, holding the removed
 * text it held and exactly the new text given it.
 */
export const placeTrackedInsertion = (
  doc: PMNode,
  from: number,
  to: number,
  pieces: readonly (readonly Mark[])[],
): number[] => {
  const stretches = new Map<Mark, { start: number; end: number }>();
  doc.nodesBetween(from, to, (node, position) => {
    if (!node.isText) {
      return true;
    }
    for (const mark of node.marks) {
      if (mark.type.name !== "comment") {
        continue;
      }
      const known =
        [...stretches.keys()].find((other) => other.eq(mark)) ?? mark;
      const start = Math.max(position, from);
      const end = Math.min(position + node.nodeSize, to);
      const stretch = stretches.get(known);
      stretches.set(
        known,
        stretch ? { start: stretch.start, end } : { start, end },
      );
    }
    return false;
  });
  // Where a piece given comments goes: next to every one of their stretches
  // (a position between two of them, past only zero-width anchors), or else
  // where the last of them ends.
  const stretchOf = (set: readonly Mark[]) => {
    const own = [...stretches].filter(([mark]) => mark.isInSet(set));
    if (own.length === 0) {
      return undefined;
    }
    const start = Math.min(...own.map(([, stretch]) => stretch.start));
    // After the removed text first, as the redline reads; before it only when
    // nothing after it is next to every comment.
    const ascending = (left: number, right: number) => left - right;
    const candidates = [
      ...own.map(([, stretch]) => stretch.end).toSorted(ascending),
      ...own.map(([, stretch]) => stretch.start).toSorted(ascending),
    ];
    for (const position of candidates) {
      const { before, after } = neighboursAt(doc, position);
      const nextToAll = own.every(
        ([mark]) =>
          Boolean(before && mark.isInSet(before.marks)) ||
          Boolean(after && mark.isInSet(after.marks)),
      );
      if (nextToAll) {
        return { start, end: position };
      }
    }
    // Disjoint comments have no position adjacent to every stretch. Put the
    // replacement before the earliest reference; accepting the deletions then
    // leaves each comment on the new text with its reference after that text.
    // The later comments reach it across removed text they did not cover, and
    // `joinCommentStretches` makes them cover it, as a saved range does.
    return { start, end: Math.min(...own.map(([, stretch]) => stretch.end)) };
  };
  const anchored = pieces.map(stretchOf);
  const positions: number[] = [];
  let previous = from;
  for (const [index, stretch] of anchored.entries()) {
    let position: number;
    if (stretch) {
      position = stretch.end;
    } else {
      const next = anchored
        .slice(index + 1)
        .find((candidate) => candidate !== undefined);
      const placedBefore = anchored
        .slice(0, index)
        .some((candidate) => candidate !== undefined);
      position = next && !placedBefore ? next.start : to;
      if (placedBefore && next) {
        position = previous;
      }
    }
    position = Math.max(position, previous);
    positions.push(position);
    previous = position;
  }
  return positions;
};

/**
 * `tr` with every comment in the paragraph at `at` covering one stretch.
 *
 * A comment is saved as two points, `w:commentRangeStart` and
 * `w:commentRangeEnd`, so it covers everything between its first and last
 * covered character. A tracked replacement over separate comments gives its
 * new text all of them, as the direct edit does, but the new text sits next
 * to only one of their removed stretches: the others then skip the removed
 * text between. The saved package cannot hold that gap and covers it, so the
 * live document covers it too. Otherwise rejecting the replacement before and
 * after a save gives different comment ranges.
 */
export const joinCommentStretches = (
  tr: Transaction,
  at: number,
): Transaction => {
  const $at = tr.doc.resolve(at);
  if (!$at.parent.inlineContent) {
    return tr;
  }
  const from = $at.start();
  const to = $at.end();
  const leaves: { node: PMNode; position: number }[] = [];
  tr.doc.nodesBetween(from, to, (node, position) => {
    if (!node.isLeaf) {
      return true;
    }
    if (!isZeroWidthAnchor(node)) {
      leaves.push({ node, position });
    }
    return false;
  });
  const comments: Mark[] = [];
  for (const { node } of leaves) {
    for (const mark of node.marks) {
      if (mark.type.name === "comment" && !mark.isInSet(comments)) {
        comments.push(mark);
      }
    }
  }
  let next = tr;
  for (const comment of comments) {
    const covered = leaves.flatMap(({ node }, index) =>
      comment.isInSet(node.marks) ? [index] : [],
    );
    const first = covered[0];
    const last = covered.at(-1);
    if (first === undefined || last === undefined) {
      continue;
    }
    for (let index = first + 1; index < last; index++) {
      const leaf = leaves[index];
      if (leaf === undefined || comment.isInSet(leaf.node.marks)) {
        continue;
      }
      next = next.addMark(
        leaf.position,
        leaf.position + leaf.node.nodeSize,
        comment,
      );
    }
  }
  return next;
};

/**
 * The comments content inserted at the block boundary `at` must carry: those
 * covering the content on both sides of it. A comment range is two points in
 * document order, so what lands between them is covered whether the range
 * crosses paragraphs, empty paragraphs or a table edge, and the saved package
 * reopens exactly so. A merge's separator lands on such a boundary too.
 */
export const commentsAcrossBlockBoundary = (doc: PMNode, at: number): Mark[] =>
  commentsAround(
    contentBesideBoundary(doc, at, -1),
    contentBesideBoundary(doc, at, 1),
  );

const withCarriedMarks = (node: PMNode, carried: readonly Mark[]): PMNode =>
  node.mark(addCarriedMarks(node.marks, carried));

/** `content` with the surveyed marks on it and the surveyed anchors around it. */
export const annotatedReplacement = ({
  annotations: { carried, leading, trailing },
  content,
}: AnnotatedReplacementOptions): AnnotatedReplacement => {
  const inline: PMNode[] = [];
  content.forEach((child) => {
    inline.push(
      carried.length === 0 ? child : withCarriedMarks(child, carried),
    );
  });

  const fragment = Fragment.fromArray([...leading, ...inline, ...trailing]);
  const leadingSize = anchorsSize(leading);
  return {
    fragment,
    leadingSize,
    contentSize: fragment.size - leadingSize - anchorsSize(trailing),
  };
};

/**
 * The marks a replacement inherits at `[from, to)`, exactly as
 * `Transaction.insertText` would compute them.
 */
export const inheritedReplacementMarks = (
  doc: PMNode,
  from: number,
  to: number,
): readonly Mark[] => {
  const $from = doc.resolve(from);
  return (
    (to === from ? $from.marks() : $from.marksAcross(doc.resolve(to))) ??
    Mark.none
  );
};
