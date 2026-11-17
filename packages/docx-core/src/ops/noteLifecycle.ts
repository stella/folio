import { Result } from "better-result";

import type { Document } from "../model/document";
import { isCommentExhausted } from "./comments";
import { combineEdits } from "./edits";
import type { ApplyOps } from "./resolve";
import { storyLifecycleEdit } from "./storyLifecycle";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "./types";

const NOTE_PART_KEYS = ["footnotes", "endnotes"] as const;
type ReplaceNotePartsOptions = {
  document: Document;
  parts: Pick<Document["package"], (typeof NOTE_PART_KEYS)[number]>;
  op: DocumentOp;
  applyOps: ApplyOps;
};

/** Note-body removal owns exhausted threads and their relationship inverses. */
export const replaceNoteParts = ({
  document,
  parts,
  op,
  applyOps,
}: ReplaceNotePartsOptions) => {
  const packageParts = { ...document.package };
  for (const key of NOTE_PART_KEYS) {
    if (Object.hasOwn(parts, key)) Reflect.set(packageParts, key, parts[key]);
    else if (key === "footnotes") delete packageParts.footnotes;
    else delete packageParts.endnotes;
  }
  const cleanup = storyLifecycleEdit(
    document,
    { ...document, package: packageParts },
    op,
    true,
  );
  if (cleanup.isErr()) return cleanup;
  const edits = [cleanup.value];
  let current = cleanup.value.document;
  const exhausted = (document.package.document.comments ?? []).filter(
    ({ id, parentId }) =>
      parentId === undefined &&
      !isCommentExhausted({ type: "document", document, id }) &&
      isCommentExhausted({ type: "document", document: current, id }),
  );
  for (const { id } of exhausted) {
    const retired = applyOps(current, [
      { type: DOCUMENT_OP_TYPES.DELETE_COMMENT, id, scope: "thread" },
    ]);
    if (retired.isErr()) return retired;
    edits.push(retired.value);
    current = retired.value.document;
  }
  return Result.ok(combineEdits(document, edits));
};
