/**
 * Whether a revision encloses an inline content control or revises its content.
 *
 * ProseMirror holds a revision as marks on leaves, and an inline control is an
 * `inline*` node rather than a leaf, so `w:ins > w:sdt` (the control was
 * inserted) and `w:sdt > w:ins` (its text was) put the same marks on the same
 * leaves. They resolve differently: rejecting the first takes the control away,
 * rejecting the second empties it (ECMA-376 §17.13.5, §17.5.2). The
 * control's `_docxEnclosingRevisionIds` attr records the first; a revision on
 * the leaves that the attr does not name revises the content.
 *
 * A revision encloses the control while it still covers everything the control
 * holds, which is also when the save leg can write it around the control.
 */

import type { Mark, MarkType, Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { expectSdtAttrs, expectTrackedChangeMarkAttrs } from "./attrs";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "./extensions/nodes/SdtExtension";

export const ENCLOSING_REVISION_IDS_ATTR = "_docxEnclosingRevisionIds";

const isInlineControl = (node: PMNode): boolean =>
  node.type.name === INLINE_CONTENT_CONTROL_NODE_NAME;

/** The revisions whose element encloses `control`. */
export const enclosingRevisionIds = (control: PMNode): readonly number[] =>
  expectSdtAttrs(control)._docxEnclosingRevisionIds ?? [];

/** `control`'s attrs with `revisionId` added to its enclosing revisions. */
export const withEnclosingRevision = (
  control: PMNode,
  revisionId: number,
): PMNode["attrs"] | null => {
  const ids = enclosingRevisionIds(control);
  return ids.includes(revisionId)
    ? null
    : { ...control.attrs, [ENCLOSING_REVISION_IDS_ATTR]: [...ids, revisionId] };
};

const REMOVED_LAYERS = {
  deletion: new Set(["deletion", "moveFrom"]),
  insertion: new Set(["insertion", "moveTo"]),
} as const;

/** Whether `mark`, or a revision it is nested in, is `revisionId` of `removeType`'s kind. */
const carriesRevision = (mark: Mark, removeType: MarkType, revisionId: number): boolean => {
  if (mark.type.name !== "insertion" && mark.type.name !== "deletion") {
    return false;
  }
  const attrs = expectTrackedChangeMarkAttrs(mark);
  if (mark.type === removeType && attrs.revisionId === revisionId) {
    return true;
  }
  const removed =
    removeType.name === "deletion" ? REMOVED_LAYERS.deletion : REMOVED_LAYERS.insertion;
  return (attrs._docxRevisionAncestors ?? []).some(
    (layer) => layer.revisionId === revisionId && removed.has(layer.type),
  );
};

const coveredBy = (node: PMNode, covers: (mark: Mark) => boolean): boolean => {
  if (node.childCount === 0) {
    return false;
  }
  for (let index = 0; index < node.childCount; index += 1) {
    const child = node.child(index);
    if (child.marks.some(covers)) {
      continue;
    }
    if (!isInlineControl(child) || !coveredBy(child, covers)) {
      return false;
    }
  }
  return true;
};

type ResolutionRemovesControlOptions = {
  control: PMNode;
  /** The revision mark type the resolution removes content under. */
  removeType: MarkType | undefined;
  /** Whether the resolution resolves revision `revisionId`. */
  resolves: (revisionId: number) => boolean;
};

/**
 * Whether resolving removes the control itself: a revision that encloses it is
 * resolved in the direction that removes its content, and still covers all of
 * it. Removing only the content leaves the control standing, emptied.
 */
export const resolutionRemovesControl = ({
  control,
  removeType,
  resolves,
}: ResolutionRemovesControlOptions): boolean => {
  if (removeType === undefined || !isInlineControl(control)) {
    return false;
  }
  // The enclosing revision is what resolves: a leaf inside it may carry a
  // revision of its own nested in it (`w:del > w:sdt > w:ins`).
  return enclosingRevisionIds(control).some(
    (revisionId) =>
      resolves(revisionId) &&
      coveredBy(control, (mark) => carriesRevision(mark, removeType, revisionId)),
  );
};

/**
 * `control`'s attrs without the enclosing revisions `resolved` names, or
 * `null` when none of them is: a resolved revision no longer exists, and an id
 * left behind would claim whatever revision is later given that id.
 */
export const withoutResolvedEnclosures = (
  control: PMNode,
  resolved: (revisionId: number) => boolean,
): PMNode["attrs"] | null => {
  const ids = enclosingRevisionIds(control);
  const remaining = ids.filter((revisionId) => !resolved(revisionId));
  if (remaining.length === ids.length) {
    return null;
  }
  return {
    ...control.attrs,
    [ENCLOSING_REVISION_IDS_ATTR]: remaining.length > 0 ? remaining : null,
  };
};

type EncloseWholeControlsOptions = {
  tr: Transaction;
  from: number;
  to: number;
  revisionId: number;
};

/**
 * Record `revisionId` as enclosing every inline control `[from, to]` covers
 * whole, opening and closing token included: the revision removes or adds the
 * control, as deleting or inserting the same range directly would. A control
 * the range covers only the content of is left alone, and its revision stays
 * one over that content.
 */
export const encloseWholeControls = ({ tr, from, to, revisionId }: EncloseWholeControlsOptions) => {
  const enclosed: { pos: number; attrs: PMNode["attrs"] }[] = [];
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!isInlineControl(node) || pos < from || pos + node.nodeSize > to) {
      return true;
    }
    const attrs = withEnclosingRevision(node, revisionId);
    if (attrs) {
      enclosed.push({ pos, attrs });
    }
    return true;
  });
  for (const { pos, attrs } of enclosed) {
    tr.setNodeMarkup(pos, undefined, attrs);
  }
};
