import type { Node as PMNode } from "prosemirror-model";

import type { BatchClaim } from "./batch-claims";
import { getOwnedRevisionRanges } from "./read";

export const EXISTING_REVISION_CONFLICT_MESSAGE =
  "The edit touches another author's pending revision.";

type RevisionGuardOptions = {
  doc: PMNode;
  author: string;
};

export const createExistingRevisionGuard = ({
  doc,
  author,
}: RevisionGuardOptions) => {
  const revisions = getOwnedRevisionRanges(doc).filter(
    (revision) => revision.author !== author,
  );
  // A zero-width edit must land strictly inside a revision carrier.
  const intersects = (from: number, to: number) =>
    revisions.some((revision) => revision.from < to && revision.to > from);
  const propertiesAt = (position: number) =>
    revisions.some(
      ({ from, to, type }) =>
        from <= position &&
        to > position &&
        type !== "insertion" &&
        type !== "deletion" &&
        type !== "formatting",
    );
  const paragraphMarkAt = (position: number) =>
    position + (doc.nodeAt(position)?.nodeSize ?? 1) - 1;
  return (claim: BatchClaim): boolean => {
    switch (claim.type) {
      case "text":
      case "annotation":
        return intersects(claim.from, claim.to);
      case "split":
        return (
          intersects(claim.from, claim.to) ||
          propertiesAt(paragraphMarkAt(claim.block))
        );
      case "rewriteBlock":
        return intersects(
          claim.block,
          claim.block + (doc.nodeAt(claim.block)?.nodeSize ?? 0),
        );
      case "paragraphProperties":
        return propertiesAt(paragraphMarkAt(claim.block));
      case "deleteBlock":
        return intersects(claim.block, claim.end);
      case "merge":
        return (
          intersects(claim.next - 1, claim.next) ||
          (claim.setsProperties &&
            propertiesAt(
              paragraphMarkAt(claim.joinsNow ? claim.block : claim.next),
            ))
        );
      case "insertion":
        return intersects(claim.at, claim.at);
      case "tableRemoval":
        return claim.ranges.some(({ from, to }) => intersects(from, to));
      case "tableStructure":
        return (
          propertiesAt(claim.table) ||
          claim.folded.some(({ from, to }) => intersects(from, to))
        );
      case "unclaimed":
        return false;
    }
  };
};
