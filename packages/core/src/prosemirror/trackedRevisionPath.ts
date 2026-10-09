import { panic } from "better-result";
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
    ...(attrs._docxResolutionJoins == null ? {} : { resolutionJoins: attrs._docxResolutionJoins }),
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

const revisionLayerOf = (mark: Mark): TrackedRevisionAncestor => {
  const attrs = expectTrackedChangeMarkAttrs(mark);
  let type: TrackedRevisionAncestor["type"];
  if (mark.type.name === "insertion") {
    type = attrs.moveKind === "moveTo" ? "moveTo" : "insertion";
  } else {
    type = attrs.moveKind === "moveFrom" ? "moveFrom" : "deletion";
  }
  return {
    type,
    revisionId: attrs.revisionId,
    author: attrs.author,
    ...(attrs.date ? { date: attrs.date } : {}),
    ...(attrs.utcDate ? { utcDate: attrs.utcDate } : {}),
    ...(attrs.initials ? { initials: attrs.initials } : {}),
    outerWrapperCount: attrs._docxOuterWrapperCount ?? 0,
  };
};

/** The authored revision layers already owned by a mark, outermost first. */
export const trackedRevisionPathOf = (mark: Mark): readonly TrackedRevisionAncestor[] => [
  ...(expectTrackedChangeMarkAttrs(mark)._docxRevisionAncestors ?? []),
  revisionLayerOf(mark),
];

/** Promote a surviving existing layer; callers cannot supply a fresh identity. */
export const retainedTrackedRevisionMark = (
  mark: Mark,
  keepLayer: (layer: TrackedRevisionAncestor) => boolean,
): Mark | null => {
  const remaining = trackedRevisionPathOf(mark).filter(keepLayer);
  const active = remaining.at(-1);
  if (!active) return null;
  const name = active.type === "insertion" || active.type === "moveTo" ? "insertion" : "deletion";
  const type = mark.type.schema.marks[name];
  if (!type) panic("A nested tracked revision has no editor mark type");
  return type.create({
    revisionId: active.revisionId,
    author: active.author,
    date: active.date ?? null,
    utcDate: active.utcDate ?? null,
    initials: active.initials ?? null,
    moveKind: active.type === "moveTo" || active.type === "moveFrom" ? active.type : null,
    _historicalFormatting: active.type === "deletion" || active.type === "moveFrom" ? true : null,
    _docxOuterWrapperCount: active.outerWrapperCount,
    _docxRevisionAncestors: remaining.length > 1 ? remaining.slice(0, -1) : null,
  });
};
