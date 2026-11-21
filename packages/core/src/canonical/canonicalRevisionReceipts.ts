import { revisionIdsIn } from "../ai-edits/read";
import type { FolioAIEditAppliedOperation } from "../ai-edits/types";
import type { Document } from "../types/document";

type StandingRevisionReceiptOptions = {
  document: Document;
  applied: ReadonlyMap<string, FolioAIEditAppliedOperation>;
};

/**
 * A receipt names only revisions the committed document holds. A later
 * operation of the batch can settle an earlier one's change away, as a run
 * patched back to the properties its pending change started from keeps none;
 * a fold into a pending change keeps that change's id, so the earlier receipt
 * still names the change it now shares.
 */
export const keepStandingRevisionReceipts = ({
  document,
  applied,
}: StandingRevisionReceiptOptions): void => {
  const standing = revisionIdsIn(document);
  for (const entry of applied.values()) {
    const { revisionIds } = entry;
    if (revisionIds === undefined) continue;
    const kept = revisionIds.filter((id) => standing.has(id));
    if (kept.length === revisionIds.length) continue;
    delete entry.revisionId;
    delete entry.revisionIds;
    const first = kept.at(0);
    if (first === undefined) continue;
    entry.revisionId = first;
    entry.revisionIds = kept;
  }
};
