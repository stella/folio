import { Result } from "better-result";

import type { Document } from "../model/document";
import type { EditorIntent, EditorIntentMode } from "./editorIntent";
import { IDENTITY_SPACES, packageIdentityKeys, slotKey } from "./ids";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp, type TextPosition } from "./types";

/** The editor compiler, for the edits one intent is made of. */
export type CompileIntent = (
  document: Document,
  intent: EditorIntent,
  mode: EditorIntentMode,
) => Result<{ ops: DocumentOp[]; selection: TextPosition }, DocumentOpRefusal>;

/**
 * The mode for the next of the edits one intent compiles to, `after` being
 * the document the edit before it left of `before`. A suggestion's next edit
 * takes for its own records the first supplied revision id that neither the
 * stamp before it nor any record that edit made holds, and keeps the rest.
 * An edit made directly takes the same mode again.
 */
export const nextCompoundMode = (
  mode: EditorIntentMode,
  before: Document,
  after: Document,
): Result<EditorIntentMode, DocumentOpRefusal> => {
  if (mode.type !== "suggesting") return Result.ok(mode);
  const known = new Set(packageIdentityKeys(before.package));
  const made = new Set(
    packageIdentityKeys(after.package).filter((key) => !known.has(key)),
  );
  const [id, ...rest] = (mode.newIds.revision ?? []).filter(
    (candidate) =>
      candidate !== mode.revision.id &&
      !made.has(slotKey({ space: IDENTITY_SPACES.REVISION, id: candidate })),
  );
  if (id === undefined)
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        message: "A compound suggestion needs additional revision ids.",
        opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      }),
    );
  return Result.ok({
    ...mode,
    revision: { ...mode.revision, id },
    newIds: { ...mode.newIds, revision: rest },
  });
};
