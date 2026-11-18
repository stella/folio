/** Source-model lock decisions shared by body and secondary-story review. */
import type { LockedControlViolation } from "@stll/docx-core/model";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  type OpStory,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import { lockViolationBetween } from "./canonicalLockGuard";

type CanonicalRevisionLocksOptions = {
  document: Document;
  story: OpStory;
  groups: readonly (readonly number[])[];
  decision: RevisionDecision;
};

export const canonicalRevisionLocks = ({
  document,
  story,
  groups,
  decision,
}: CanonicalRevisionLocksOptions) => {
  const held: { ids: readonly number[]; violation: LockedControlViolation }[] =
    [];
  const unlocked: (readonly number[])[] = [];
  for (const ids of groups) {
    const alone = applyDocumentOps(document, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story,
        revisionIds: ids,
        decision,
      },
    ]);
    // A non-lock refusal is propagated by staging the requested resolution.
    const violation = alone.isErr()
      ? null
      : lockViolationBetween(document, alone.value.document);
    if (violation === null) unlocked.push(ids);
    else held.push({ ids, violation });
  }
  return { held, unlocked };
};
