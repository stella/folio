import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { TRACKED_RUN_INLINE_ATOM_DISPOSITIONS } from "./trackedRunInlineAtoms";

/**
 * Whether an inline atom extends a comment range it sits in.
 *
 * - `extends`: a comment mark on the atom keeps the comment's range alive.
 * - `inert`: the atom is invisible to the reader, so a comment mark on it would
 *   be a ghost range nobody can see or select. Comment integrity strips the mark
 *   and never counts the atom as range content.
 *
 * This is separate from revision ownership: a cached rendered page break must
 * stay inside its tracked-change wrapper, but it can never hold a comment on
 * its own.
 */
export type CommentRangeInlineAtomParticipation = "extends" | "inert";

/** Keyed by the tracked-run atom set, which the fixed-point test binds to the schema. */
export const COMMENT_RANGE_INLINE_ATOM_PARTICIPATION: Readonly<{
  readonly bookmarkBoundary: "extends";
  readonly commentReference: "inert";
  readonly commentRangeBoundary: "inert";
  readonly field: "extends";
  readonly hardBreak: "extends";
  readonly image: "extends";
  readonly math: "extends";
  readonly moveRangeBoundary: "extends";
  readonly noteMarker: "extends";
  readonly pageBreakRun: "extends";
  readonly preservedXml: "extends";
  readonly rangeAnchor: "inert";
  readonly renderedPageBreak: "inert";
  readonly shape: "extends";
  readonly structuredField: "extends";
  readonly symbol: "extends";
  readonly tab: "extends";
  readonly text: "extends";
  readonly textBoxAnchor: "extends";
}> = Object.freeze({
  bookmarkBoundary: "extends",
  commentReference: "inert",
  commentRangeBoundary: "inert",
  field: "extends",
  hardBreak: "extends",
  image: "extends",
  math: "extends",
  moveRangeBoundary: "extends",
  noteMarker: "extends",
  pageBreakRun: "extends",
  preservedXml: "extends",
  rangeAnchor: "inert",
  renderedPageBreak: "inert",
  shape: "extends",
  structuredField: "extends",
  symbol: "extends",
  tab: "extends",
  text: "extends",
  textBoxAnchor: "extends",
} as const satisfies Readonly<
  Record<
    keyof typeof TRACKED_RUN_INLINE_ATOM_DISPOSITIONS,
    CommentRangeInlineAtomParticipation
  >
>);

const participationByName = new Map<
  string,
  CommentRangeInlineAtomParticipation
>(Object.entries(COMMENT_RANGE_INLINE_ATOM_PARTICIPATION));

/** Whether a node is an inline atom that must never carry or extend a comment range. */
export const isCommentRangeInertAtom = (node: PMNode): boolean => {
  if (!node.isInline || !node.isAtom) {
    return false;
  }
  const participation = participationByName.get(node.type.name);
  if (!participation) {
    panic(
      `Inline atom ${JSON.stringify(node.type.name)} has no comment-range participation`,
    );
  }
  return participation === "inert";
};
