import { MAX_REVISION_ID } from "../model/document";
import { IDENTITY_SPACES, slotKey } from "./ids";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { DocumentOp, RevisionStamp } from "./types";

/**
 * Whether a tracked operation's stamp can name the first record it creates:
 * a revision id in range that no record in the package carries.
 */
export const trackedStampRefusal = (
  op: DocumentOp,
  stamp: RevisionStamp,
  hasRecord: (key: string) => boolean,
): DocumentOpRefusal | undefined => {
  // A revision id reserves no value: every integer in range names a record.
  const revisionId = stamp.id;
  if (
    !Number.isInteger(revisionId) ||
    revisionId < 0 ||
    revisionId > MAX_REVISION_ID
  ) {
    return new DocumentOpRefusal({
      opType: op.type,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      message: `${revisionId} cannot be a revision id.`,
    });
  }
  const key = slotKey({ space: IDENTITY_SPACES.REVISION, id: stamp.id });
  if (hasRecord(key)) {
    return new DocumentOpRefusal({
      opType: op.type,
      reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      message: `Revision id ${stamp.id} is already used in the package.`,
    });
  }
  return undefined;
};
