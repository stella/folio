/** A note goes with its reference: tracking its last live reference's deletion deletes its text too. */
import { Result } from "better-result";

import type { Document, Paragraph } from "../model/document";
import { deletedContent, rangeEndIdsIn } from "./blockDeletion";
import { storyParagraphs } from "./blocks";
import { combineEdits, type DocumentEdit } from "./edits";
import { freshenIdentities } from "./identity";
import {
  IDENTITY_SPACES,
  identityKeysIn,
  packageIdentityKeyIndex,
  packageIdentityKeys,
  slotKey,
} from "./ids";
import { notesLosingLiveReferences } from "./noteMarks";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import { carriesStamp, stampInfo } from "./review";
import { findStoryBody } from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type NoteStory,
  type RevisionStamp,
} from "./types";

type NoteDeletionOptions = {
  document: Document;
  story: NoteStory;
  stamp: RevisionStamp;
  revisionIds: readonly number[];
  op: DocumentOp;
  applyOps: ApplyOps;
};

/**
 * Every paragraph of a note deleted under `stamp`: its content wrapped in a
 * deletion, and its mark deleted where it records no other change. Each
 * record takes one of `revisionIds`: `stamp`'s own id is the reference's
 * deletion's, which none of them may share. `undefined` when the note holds
 * nothing a deletion could record, or what a tracked deletion cannot hold.
 */
const deleteNoteContent = ({
  document,
  story,
  stamp,
  revisionIds,
  op,
  applyOps,
}: NoteDeletionOptions):
  Result<DocumentEdit, DocumentOpRefusal> | undefined => {
  const body = findStoryBody(document, story);
  if (body === undefined) return undefined;
  const rangedComments = rangeEndIdsIn(body.content);
  const before: Paragraph[] = [];
  const after: Paragraph[] = [];
  for (const { paragraph } of storyParagraphs(body)) {
    const form = deletedContent(paragraph, stamp, rangedComments);
    if (form.kind === "refused") return undefined;
    let next =
      form.kind === "deleted"
        ? { ...paragraph, content: form.content }
        : paragraph;
    if (next.pPrMark === undefined)
      next = { ...next, pPrMark: { kind: "del", info: stampInfo(stamp) } };
    if (next === paragraph) continue;
    before.push(paragraph);
    after.push(next);
  }
  if (after.length === 0) return undefined;
  const beforeKeys = new Set(identityKeysIn(before));
  const fresh = freshenIdentities({
    before,
    after,
    newIds: { revision: [...revisionIds] },
    usedElsewhere: () =>
      new Set(
        packageIdentityKeys(document.package).filter(
          (key) => !beforeKeys.has(key),
        ),
      ),
  });
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  switch (fresh.kind) {
    case "needsIds":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        "Deleting a note with its reference needs more revision ids.",
      );
    case "invalidId":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        "Deleting a note with its reference received an invalid revision id.",
      );
    case "fresh":
      return applyOps(
        document,
        fresh.paragraphs.map((paragraph, index) => ({
          type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
          story,
          expected: [before[index] ?? paragraph],
          blocks: [paragraph],
        })),
      );
    default: {
      const unreachable: never = fresh;
      return unreachable;
    }
  }
};

type DeleteNotesWithReferencesOptions = {
  before: Document;
  edit: DocumentEdit;
  op: DocumentOp;
  stamp: RevisionStamp;
  applyOps: ApplyOps;
};

/**
 * A tracked operation that deletes the last live reference to a note, in
 * any story, deletes the note's text and paragraph marks with it, by the
 * reference deletion's author at its date, as the reference implementation
 * records it: rejecting that deletion gives the note its text back, and
 * accepting it removes the note. A suggestion stays its author's until it
 * is accepted, so it deletes no note. The records take the operation's
 * unused new revision ids.
 */
export const deleteNotesWithReferences = ({
  before,
  edit,
  op,
  stamp,
  applyOps,
}: DeleteNotesWithReferencesOptions): Result<
  DocumentEdit,
  DocumentOpRefusal
> => {
  if (stamp.suggestionId !== undefined) return Result.ok(edit);
  const notes = notesLosingLiveReferences(before, edit.document).filter(
    ({ deletion }) => carriesStamp(deletion.info, stamp),
  );
  if (notes.length === 0) return Result.ok(edit);
  const supplied = "newIds" in op ? (op.newIds?.revision ?? []) : [];
  const edits = [edit];
  let current = edit.document;
  for (const { story, deletion } of notes) {
    const used = packageIdentityKeyIndex(current.package);
    const deleted = deleteNoteContent({
      document: current,
      story,
      stamp: { ...stamp, id: deletion.info.id },
      revisionIds: supplied.filter(
        (id) =>
          id !== stamp.id &&
          !used.has(slotKey({ space: IDENTITY_SPACES.REVISION, id })),
      ),
      op,
      applyOps,
    });
    if (deleted === undefined) continue;
    if (deleted.isErr()) return deleted;
    edits.push(deleted.value);
    current = deleted.value.document;
  }
  return Result.ok(edits.length === 1 ? edit : combineEdits(before, edits));
};
