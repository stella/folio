import { panic } from "better-result";
import type { Mark, Node as PMNode } from "prosemirror-model";
import { type Step, type StepMap, Transform } from "prosemirror-transform";

import { isCommentRangeInertAtom } from "./commentRangeInlineAtoms";
import { readCommentReferenceAttrs } from "./commentReferenceAttrs";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";

/**
 * What a document must hold for every comment: exactly one `commentReference`
 * node, after the last position its `comment` mark covers.
 *
 * The reference is the run Word paints the visible comment mark with, so a
 * comment without one is invisible in Word and a duplicate paints the mark
 * twice. Editing cannot be trusted to keep that true — a paste duplicates a
 * reference, deleting the commented text orphans one, and a command that adds
 * a comment mark creates none — so the repair is computed here and applied by
 * the editor extension and shared whole-story revision resolver.
 */

type PositionedReference = {
  position: number;
  /** Node size, so a caller can delete the node without re-reading it. */
  size: number;
};

export type CommentReferenceOccurrence =
  | (PositionedReference & { status: "valid"; commentId: number })
  | (PositionedReference & { status: "malformed" });

/** Where a comment's mark last covers content, and how far that content runs. */
export type CommentMarkExtent = {
  commentId: number;
  /** Document position just past the last inline node carrying the mark. */
  end: number;
};

export type CommentReferenceRepair =
  | { type: "delete"; position: number; size: number }
  | { type: "insert"; position: number; commentId: number };

type CommentReferenceRepairInput = {
  references: readonly CommentReferenceOccurrence[];
  markExtents: readonly CommentMarkExtent[];
};

export const COMMENT_REFERENCE_NODE_NAME = "commentReference";

type CommentReferenceSurvey = {
  references: CommentReferenceOccurrence[];
  markExtents: CommentMarkExtent[];
  unanchoredMarks: { position: number; size: number; mark: Mark }[];
};

export const surveyCommentReferences = (
  doc: PMNode,
): CommentReferenceSurvey => {
  const references: CommentReferenceOccurrence[] = [];
  const markEnds = new Map<number, number>();
  const unanchoredMarks: CommentReferenceSurvey["unanchoredMarks"] = [];

  doc.descendants((node, position) => {
    const inert = isCommentRangeInertAtom(node);
    if (inert) {
      for (const mark of node.marks) {
        if (mark.type.name === "comment")
          unanchoredMarks.push({ position, size: node.nodeSize, mark });
      }
    }
    if (node.type.name === COMMENT_REFERENCE_NODE_NAME) {
      const attrs = readCommentReferenceAttrs(node);
      references.push(
        attrs.ok
          ? {
              status: "valid",
              commentId: attrs.value.commentId,
              position,
              size: node.nodeSize,
            }
          : { status: "malformed", position, size: node.nodeSize },
      );
      return false;
    }
    // Point comments have no marked content, so their range anchor is the
    // surviving extent that keeps the reference.
    if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
      const attrs = readRangeAnchorAttrs(node);
      if (attrs.ok && attrs.value.start.type === "commentRangeStart") {
        const { id } = attrs.value.start;
        markEnds.set(
          id,
          Math.max(markEnds.get(id) ?? 0, position + node.nodeSize),
        );
      }
      return false;
    }
    if (inert) return false;
    if (!node.isInline) {
      return true;
    }
    for (const mark of node.marks) {
      if (mark.type.name !== "comment") {
        continue;
      }
      const commentId = mark.attrs["commentId"];
      if (typeof commentId !== "number") {
        continue;
      }
      markEnds.set(
        commentId,
        Math.max(markEnds.get(commentId) ?? 0, position + node.nodeSize),
      );
    }
    return true;
  });

  return {
    references,
    unanchoredMarks,
    markExtents: [...markEnds].map(([commentId, end]) => ({ commentId, end })),
  };
};
/**
 * The repairs that make a document total, as positions in the document they
 * were read from. Deletions come first, last-first, so they apply in sequence
 * without mapping; the insertions that follow still carry pre-repair positions
 * and the applier maps them through the transaction.
 */
export const planCommentReferenceRepairs = ({
  references,
  markExtents,
}: CommentReferenceRepairInput): readonly CommentReferenceRepair[] => {
  const markedIds = new Set(markExtents.map(({ commentId }) => commentId));
  const deletions: CommentReferenceRepair[] = [];
  const keptIds = new Set<number>();

  for (const reference of references) {
    if (reference.status === "malformed") {
      deletions.push({
        type: "delete",
        position: reference.position,
        size: reference.size,
      });
      continue;
    }
    // An orphan: the comment it names no longer covers anything, so nothing
    // in the editor can show it. A second reference for a live comment is a
    // duplicate; the first one wins so the repair is order-independent.
    if (!markedIds.has(reference.commentId)) {
      deletions.push({
        type: "delete",
        position: reference.position,
        size: reference.size,
      });
      continue;
    }
    if (keptIds.has(reference.commentId)) {
      deletions.push({
        type: "delete",
        position: reference.position,
        size: reference.size,
      });
      continue;
    }
    keptIds.add(reference.commentId);
  }

  const insertions: CommentReferenceRepair[] = [];
  for (const { commentId, end } of markExtents) {
    if (keptIds.has(commentId)) {
      continue;
    }
    insertions.push({ type: "insert", position: end, commentId });
  }

  return [
    ...deletions.toSorted((first, second) => second.position - first.position),
    ...insertions.toSorted((first, second) => second.position - first.position),
  ];
};

type CommentReferenceResolution = {
  doc: PMNode;
  steps: Step[];
  positionMaps: StepMap[];
};

export const resolveCommentReferences = (
  doc: PMNode,
): CommentReferenceResolution | null => {
  const survey = surveyCommentReferences(doc);
  const repairs = planCommentReferenceRepairs(survey);
  if (repairs.length === 0 && survey.unanchoredMarks.length === 0) return null;

  const transform = new Transform(doc);
  for (const { position, size, mark } of survey.unanchoredMarks)
    transform.removeMark(position, position + size, mark);
  const referenceType = doc.type.schema.nodes[COMMENT_REFERENCE_NODE_NAME];
  if (!referenceType && repairs.some(({ type }) => type === "insert"))
    return panic("Comment reference repairs require a commentReference node.");
  for (const repair of repairs) {
    if (repair.type === "delete") {
      transform.delete(repair.position, repair.position + repair.size);
      continue;
    }
    transform.insert(
      transform.mapping.map(repair.position),
      referenceType.create({ commentId: repair.commentId }),
    );
  }
  return {
    doc: transform.doc,
    steps: transform.steps,
    positionMaps: transform.mapping.maps,
  };
};
