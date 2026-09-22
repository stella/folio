/**
 * What a document must hold for every comment: exactly one `commentReference`
 * node, after the last position its `comment` mark covers.
 *
 * The reference is the run Word paints the visible comment mark with, so a
 * comment without one is invisible in Word and a duplicate paints the mark
 * twice. Editing cannot be trusted to keep that true — a paste duplicates a
 * reference, deleting the commented text orphans one, and a command that adds
 * a comment mark creates none — so the repair is computed here and applied by
 * the extension's `appendTransaction`, rather than left to each call site.
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
      deletions.push({ type: "delete", position: reference.position, size: reference.size });
      continue;
    }
    // An orphan: the comment it names no longer covers anything, so nothing
    // in the editor can show it. A second reference for a live comment is a
    // duplicate; the first one wins so the repair is order-independent.
    if (!markedIds.has(reference.commentId)) {
      deletions.push({ type: "delete", position: reference.position, size: reference.size });
      continue;
    }
    if (keptIds.has(reference.commentId)) {
      deletions.push({ type: "delete", position: reference.position, size: reference.size });
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
