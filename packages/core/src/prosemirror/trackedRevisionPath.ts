import type { Mark, Node as PMNode } from "prosemirror-model";

import { expectInlineWrapperMarkAttrs, expectTrackedChangeMarkAttrs } from "./attrs";
import type { InlineWrapperLayer, TrackedRevisionAncestor } from "./schema/marks";
import { INLINE_WRAPPER_MARK_NAME } from "./extensions/marks/InlineWrapperExtension";

/** The wrappers `node` sits inside, outermost first; empty when it sits in none. */
export const inlineWrapperStackOf = (node: PMNode): readonly InlineWrapperLayer[] => {
  const mark = node.marks.find((candidate) => candidate.type.name === INLINE_WRAPPER_MARK_NAME);
  return mark ? expectInlineWrapperMarkAttrs(mark).stack : [];
};

/** One layer of the path used to write and enumerate a tracked run. */
export const trackedRevisionLayerOf = (mark: Mark, node: PMNode): TrackedRevisionAncestor => {
  const attrs = expectTrackedChangeMarkAttrs(mark);
  const type = mark.type.name === "insertion" ? "insertion" : "deletion";
  const moveType = type === "insertion" ? "moveTo" : "moveFrom";
  return {
    type: attrs.moveKind === moveType ? moveType : type,
    revisionId: attrs.revisionId,
    author: attrs.author || "Unknown",
    ...(attrs.date ? { date: attrs.date } : {}),
    ...(attrs.utcDate ? { utcDate: attrs.utcDate } : {}),
    ...(attrs.initials ? { initials: attrs.initials } : {}),
    outerWrapperCount: attrs._docxOuterWrapperCount ?? inlineWrapperStackOf(node).length,
  };
};

type EnclosingInsertionAncestorsOptions = {
  /** The insertion a deletion mark shares its text with, if any. */
  insertionMark: Mark | undefined;
  ancestors: readonly TrackedRevisionAncestor[];
  node: PMNode;
};

/** The deletion's ancestor path, with the insertion it sits in as the outermost layer. */
export const enclosingInsertionAncestors = ({
  insertionMark,
  ancestors,
  node,
}: EnclosingInsertionAncestorsOptions): readonly TrackedRevisionAncestor[] => {
  if (!insertionMark) {
    return ancestors;
  }
  const attrs = expectTrackedChangeMarkAttrs(insertionMark);
  if (ancestors.some(({ revisionId }) => revisionId === attrs.revisionId)) {
    return ancestors;
  }
  const layer = trackedRevisionLayerOf(insertionMark, node);
  return [layer, ...ancestors];
};
