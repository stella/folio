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
 * The alignment helper preserves a comment's coverage. Unchanged text
 * keeps its ranges; inserted text takes the comments of the replaced aligned
 * span. It can propose a point at the replacement start for an emptied range;
 * downstream ranged-comment exhaustion may retire that comment. The applier
 * decides alignment once for direct and tracked edits.
 * Canonical replacements accept only coverage the source anchors already
 * express. Pure deletion follows ranged-comment exhaustion; unproven point
 * transitions refuse before mutation.
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
 * Removing the owning block removes its comment. Point anchors proposed by
 * the alignment helper do not override ranged-comment retirement; the
 * canonical guard refuses point transitions whose outcome is not proven.
 */

import { panic } from "better-result";
import {
  Fragment,
  Mark,
  type Node as PMNode,
  type Schema,
} from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { isCommentExhausted } from "@stll/docx-core/ops";

import {
  commentAnchorIdsForNode,
  isCommentAnchorCarrier,
} from "../ai-edits/comment-lifecycle";
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
  readonly hyphenElement: "followsTheText";
  readonly insertion: "followsTheText";
  readonly deletion: "followsTheText";
  readonly runPropertyChange: "followsTheText";
  readonly inlineWrapper: "followsTheText";
} = {
  comment: "carry",
  hyperlink: "carry",
  footnoteRef: "followsTheText",
  runIdentity: "followsTheText",
  hyphenElement: "followsTheText",
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
    markers: [
      { type: "commentRangeStart", id },
      { type: "commentRangeEnd", id },
    ],
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
      if (attrs.ok)
        for (const marker of attrs.value.markers)
          if (marker.type === "commentRangeStart") retained.add(marker.id);
    }
    if (!node.isInline || isCommentAnchorCarrier(node)) return true;
    const removedWithStructure = removedStructures.some(
      ({ from, to }) => position > from && position < to,
    );
    for (const mark of node.marks) {
      if (mark.type.name !== "comment") continue;
      // A point comment lives as long as its reference, not its display text.
      if (mark.attrs["point"] === true) continue;
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
 * Which content of a replaced span lends its marks to the replacement.
 * `asWritten`: every leaf. `asAccepted`: the leaves an earlier edit has not
 * already deleted. A pending deletion's text is gone once accepted and is
 * not what a direct edit replaces, so its comments and links stay on it.
 */
type ReplacedContent = "asWritten" | "asAccepted";

const lendsMarks = (node: PMNode, content: ReplacedContent): boolean => {
  switch (content) {
    case "asWritten":
      return true;
    case "asAccepted":
      return !node.marks.some(({ type }) => type.name === "deletion");
    default: {
      const unhandled: never = content;
      throw new Error(`Unhandled replaced content ${String(unhandled)}`);
    }
  }
};

type SurveyOptions = {
  from: number;
  to: number;
  content: ReplacedContent;
  containers: "preserve" | "descend";
};

const surveyAnnotations = (
  doc: PMNode,
  { from, to, content, containers }: SurveyOptions,
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
    for (const mark of lendsMarks(node, content) ? node.marks : []) {
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
      return (
        containers === "descend" && !node.isLeaf && lendsMarks(node, content)
      );
    }
    (opensRange(node, placement) ? leading : trailing).push(node);
    return false;
  });

  return { carried, leading, trailing };
};

/**
 * The annotations carried by `[from, to)`. An anchor counts only when the
 * range covers it whole: a partially covered atom is not deleted by the
 * replacement either.
 */
export const surveyReplacedAnnotations = (
  doc: PMNode,
  from: number,
  to: number,
): ReplacedAnnotations =>
  surveyAnnotations(doc, {
    from,
    to,
    content: "asWritten",
    containers: "preserve",
  });

/**
 * {@link surveyReplacedAnnotations} over the span as it reads once accepted:
 * what a replacement of it carries, which a direct edit of the same text
 * carries too.
 */
export const surveyAcceptedReplacedAnnotations = (
  doc: PMNode,
  from: number,
  to: number,
): ReplacedAnnotations =>
  surveyAnnotations(doc, {
    from,
    to,
    content: "asAccepted",
    containers: "preserve",
  });

type ReplacementCommentAlignmentOptions = {
  doc: PMNode;
  source: { from: number; to: number };
  retiredCommentIds?: ReadonlySet<number> | undefined;
  /**
   * The replacement is tracked: the words it deletes keep their comment
   * ranges, so rejecting gives every range back and accepting drops a
   * comment whose words all went. Only a point comment or a selected marker
   * still needs the comment to move.
   */
  deletionsKeepComments?: boolean;
  changes: readonly {
    from: number;
    to: number;
    text: string;
    markers: "retained" | "selected";
  }[];
};

/**
 * Whether source boundaries already express the replacement's aligned ranges.
 * A changed stretch lends no comment to its replacement: the new words go
 * past the end of every range the stretch reaches. Pure deletions within surviving source content follow the
 * ranged-comment exhaustion owner; other collapses need a point policy.
 * Decide over the whole replacement, before any of its fragments are applied.
 */
export const preservesReplacementCommentAlignment = ({
  doc,
  source,
  changes,
  retiredCommentIds,
  deletionsKeepComments = false,
}: ReplacementCommentAlignmentOptions): boolean => {
  let previousEnd = source.from;
  for (const change of changes) {
    if (
      change.from < previousEnd ||
      change.to < change.from ||
      change.to > source.to
    )
      panic(
        "Replacement changes must be ordered inside their source coverage.",
      );
    previousEnd = change.to;
  }
  const collapsing: Mark[] = [];
  const continued: Mark[] = [];
  for (const change of changes) {
    const annotations = surveyAnnotations(doc, {
      from: change.from,
      to: change.to,
      content: "asAccepted",
      containers: "descend",
    });
    const comments = annotations.carried.filter(
      (mark) =>
        mark.type.name === "comment" &&
        !retiredCommentIds?.has(mark.attrs["commentId"]),
    );
    if (comments.some((mark) => mark.attrs["point"] === true)) return false;
    if (change.markers === "selected" && comments.length > 0) return false;
    if (deletionsKeepComments) continue;
    let internalCarrier = false;
    doc.nodesBetween(change.from, change.to, (node, position) => {
      if (
        isCommentAnchorCarrier(node) &&
        change.from < position &&
        position < change.to
      ) {
        const ids = commentAnchorIdsForNode(node);
        if (ids.length === 0 || ids.some((id) => !retiredCommentIds?.has(id)))
          internalCarrier = true;
      }
      return !node.isLeaf && lendsMarks(node, "asAccepted");
    });
    if (internalCarrier) return false;
    if (change.text === "") {
      for (const mark of comments)
        if (!mark.isInSet(collapsing)) collapsing.push(mark);
      continue;
    }
    // The new words take no comment from the words they replace: they go past
    // the end of a range the change reaches, and one it covers wholly goes
    // with its words. A range starting in the change keeps them, as accepting
    // the tracked replacement leaves them in it.
    for (const mark of comments)
      if (!mark.isInSet(continued)) continued.push(mark);
  }
  const candidates = collapsing.filter((mark) => !mark.isInSet(continued));
  if (candidates.length === 0) return true;
  const retained: Mark[] = [];
  let sourceContentSurvives = false;
  let changeIndex = 0;
  doc.descendants((node, position) => {
    if (!node.isInline) return true;
    if (!node.isLeaf) return lendsMarks(node, "asAccepted");
    if (isZeroWidthAnchor(node) || !lendsMarks(node, "asAccepted"))
      return false;
    const from = position;
    const to = position + node.nodeSize;
    const sourceFrom = Math.max(source.from, from);
    const sourceTo = Math.min(source.to, to);
    while ((changes[changeIndex]?.to ?? Infinity) <= from) changeIndex++;
    let removed = 0;
    let sourceRemoved = 0;
    for (let index = changeIndex; index < changes.length; index++) {
      const change = changes[index];
      if (change === undefined || change.from >= to) break;
      removed += Math.max(
        0,
        Math.min(to, change.to) - Math.max(from, change.from),
      );
      sourceRemoved += Math.max(
        0,
        Math.min(sourceTo, change.to) - Math.max(sourceFrom, change.from),
      );
    }
    if (sourceFrom < sourceTo && sourceRemoved < sourceTo - sourceFrom)
      sourceContentSurvives = true;
    if (removed < to - from)
      for (const mark of candidates)
        if (mark.isInSet(node.marks) && !mark.isInSet(retained))
          retained.push(mark);
    return false;
  });
  const deletesOnly = changes.every(({ text }) => text === "");
  return candidates.every((mark) => {
    const survivingNonCarrierContent = mark.isInSet(retained);
    // Candidates are non-point marks. The whole-document scan proves each
    // range's remaining coverage, including coverage outside the edited source.
    return (
      survivingNonCarrierContent ||
      (deletesOnly &&
        sourceContentSurvives &&
        isCommentExhausted({
          type: "coverage",
          anchorKind: "range",
          survivingNonCarrierContent,
        }))
    );
  });
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

/**
 * The content before `from` and after `to`, as a saved range's two points
 * read it: in their paragraph past zero-width anchors, and past the
 * paragraph's edge, empty paragraphs and table edges when nothing is left in
 * it on that side.
 */
const contentAround = (
  doc: PMNode,
  from: number,
  to: number = from,
): { before: PMNode | null; after: PMNode | null } => {
  const { before } = neighboursAt(doc, from);
  const { after } = neighboursAt(doc, to);
  const $from = doc.resolve(from);
  const $to = doc.resolve(to);
  return {
    before:
      before ??
      ($from.parent.inlineContent && $from.depth > 0
        ? contentBesideBoundary(doc, $from.before(), -1)
        : null),
    after:
      after ??
      ($to.parent.inlineContent && $to.depth > 0
        ? contentBesideBoundary(doc, $to.after(), 1)
        : null),
  };
};

/** The comments on both `before` and `after`: text between them must carry them. */
const commentsAround = (before: PMNode | null, after: PMNode | null): Mark[] =>
  before && after
    ? before.marks.filter(
        (mark) => mark.type.name === "comment" && mark.isInSet(after.marks),
      )
    : [];

/**
 * The comments covering the content on both sides of `[from, to)`, across a
 * paragraph's edge too: text an edit puts there lies between the two points
 * a save writes for each, so the saved package reopens with it covered.
 */
export const commentsSpanningInsertion = (
  doc: PMNode,
  from: number,
  to: number = from,
): Mark[] => {
  const { before, after } = contentAround(doc, from, to);
  return commentsAround(before, after);
};

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
  // At a paragraph's edge the neighbour is across it: a comment covering the
  // content on both sides covers what lands between, as its saved range does.
  const { before, after } = contentAround(doc, at);
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
