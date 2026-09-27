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
 * Word moves a range start that sits inside a replaced span to the beginning
 * of the replacement and lets the range keep covering the new text; it keeps
 * a link across a replacement of its display text. A range wholly inside the
 * span collapses onto the replacement, and disappears only when every
 * character it covered is gone. One rule per marker kind, and that is the
 * whole contract:
 *
 * - `comment` mark: the replacement carries every comment the span covered, so
 *   a range that started inside it starts at the replacement instead.
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
 * Deleting the whole of a comment's span leaves it no mark anywhere, so
 * `CommentReferenceExtension` drops the orphan reference and the save drops
 * the comment: the one case where a range is meant to disappear.
 */

import { Fragment, Mark, type Node as PMNode, type Schema } from "prosemirror-model";

import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { COMMENT_REFERENCE_NODE_NAME } from "./extensions/nodes/CommentReferenceExtension";
import { MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/MoveRangeBoundaryExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import { TEXT_BOX_ANCHOR_NODE_NAME } from "./extensions/nodes/TextBoxAnchorExtension";
import { readMoveRangeBoundaryAttrs } from "./moveRangeBoundaryAttrs";
import { isZeroWidthAnchor } from "./zeroWidthAnchors";

/**
 * What a replacement does with each mark the schema declares
 * `inclusive: false`. Boundary inheritance already handles every other mark,
 * so this table is exactly the set that needs a decision; a new
 * non-inclusive mark that is missing from it fails
 * `replacedAnnotations.test.ts` rather than silently following the text.
 */
export const NON_INCLUSIVE_MARK_DISPOSITION = {
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
  NON_INCLUSIVE_MARK_DISPOSITION[mark.type.name as keyof typeof NON_INCLUSIVE_MARK_DISPOSITION] ===
  "carry";

/** Anchors an edit must put back, and where relative to the replacement. */
const ANCHOR_PLACEMENT = {
  [COMMENT_REFERENCE_NODE_NAME]: "after",
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

const NOTHING_REPLACED: ReplacedAnnotations = { carried: [], leading: [], trailing: [] };

export const hasReplacedAnnotations = ({
  carried,
  leading,
  trailing,
}: ReplacedAnnotations): boolean => carried.length > 0 || leading.length > 0 || trailing.length > 0;

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
    if (placement === undefined || position < from || position + node.nodeSize > to) {
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

/** The first (`1`) or last (`-1`) child of `paragraph` that is not a zero-width anchor. */
const edgeContent = (paragraph: PMNode, direction: -1 | 1): PMNode | null => {
  for (
    let index = direction === 1 ? 0 : paragraph.childCount - 1;
    index >= 0 && index < paragraph.childCount;
    index += direction
  ) {
    const child = paragraph.child(index);
    if (!isZeroWidthAnchor(child)) {
      return child;
    }
  }
  return null;
};

/**
 * The content either side of `at` in its paragraph, past zero-width anchors
 * (a comment's own reference sits between its text and what follows).
 */
const neighboursAt = (doc: PMNode, at: number): { before: PMNode | null; after: PMNode | null } => {
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
    ? before.marks.filter((mark) => mark.type.name === "comment" && mark.isInSet(after.marks))
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
  const sets: (readonly Mark[])[] = [before?.marks ?? [], ...pieces, after?.marks ?? []];
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
    const covered = sets.flatMap((set, index) => (comment.isInSet(set) ? [index] : []));
    const first = covered[0]!;
    const last = covered.at(-1)!;
    // Only a stretch that reaches a neighbour continues the comment.
    if (first !== 0 && last !== sets.length - 1) {
      continue;
    }
    for (let index = Math.max(first, 1); index <= Math.min(last, pieces.length); index++) {
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
      const known = [...stretches.keys()].find((other) => other.eq(mark)) ?? mark;
      const start = Math.max(position, from);
      const end = Math.min(position + node.nodeSize, to);
      const stretch = stretches.get(known);
      stretches.set(known, stretch ? { start: stretch.start, end } : { start, end });
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
    return { start, end: Math.max(...own.map(([, stretch]) => stretch.end)) };
  };
  const anchored = pieces.map(stretchOf);
  const positions: number[] = [];
  let previous = from;
  for (const [index, stretch] of anchored.entries()) {
    let position: number;
    if (stretch) {
      position = stretch.end;
    } else {
      const next = anchored.slice(index + 1).find((candidate) => candidate !== undefined);
      const placedBefore = anchored.slice(0, index).some((candidate) => candidate !== undefined);
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
 * The comments text inserted between the end of `first` and the start of
 * `next` must carry: those covering both sides. A merge's separator lands
 * there.
 */
export const commentsAcrossParagraphBoundary = (first: PMNode, next: PMNode): Mark[] =>
  commentsAround(edgeContent(first, -1), edgeContent(next, 1));

const withCarriedMarks = (node: PMNode, carried: readonly Mark[]): PMNode =>
  node.mark(addCarriedMarks(node.marks, carried));

/** `content` with the surveyed marks on it and the surveyed anchors around it. */
export const annotatedReplacement = ({
  annotations: { carried, leading, trailing },
  content,
}: AnnotatedReplacementOptions): AnnotatedReplacement => {
  const inline: PMNode[] = [];
  content.forEach((child) => {
    inline.push(carried.length === 0 ? child : withCarriedMarks(child, carried));
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
  return (to === from ? $from.marks() : $from.marksAcross(doc.resolve(to))) ?? Mark.none;
};
