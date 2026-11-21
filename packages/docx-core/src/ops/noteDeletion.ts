/** A note goes with its reference: tracking its last live reference's deletion deletes its text too. */
import { Result } from "better-result";

import type { Document, Paragraph } from "../model/document";
import { deletedContent, rangeEndIdsIn } from "./blockDeletion";
import { storyParagraphs } from "./blocks";
import { freshenIdentities } from "./identity";
import {
  IDENTITY_SPACES,
  identityKeysIn,
  packageIdentityKeyIndex,
  slotKey,
} from "./ids";
import { notesLosingLiveReferences } from "./noteMarks";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { carriesStamp, stampInfo } from "./review";
import { findStoryBody } from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type NoteStory,
  type RevisionStamp,
} from "./types";

type NoteContentDeletion = { before: Paragraph[]; after: Paragraph[] };

/**
 * Every paragraph of a note deleted under `stamp`: its content wrapped in a
 * deletion, and its mark deleted where it records no other change.
 * `undefined` when the note holds nothing a deletion could record, or what a
 * tracked deletion cannot hold.
 */
const deletedNoteContent = (
  document: Document,
  story: NoteStory,
  stamp: RevisionStamp,
): NoteContentDeletion | undefined => {
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
  return after.length === 0 ? undefined : { before, after };
};

type NoteDeletionOpsOptions = {
  /** The document before the tracked edit. */
  before: Document;
  /** The document the tracked edit left. */
  after: Document;
  /** The edit's stamp: the notes whose last live reference it deleted follow it. */
  stamp: RevisionStamp;
  /** Ids the note records may take; those `after` holds are passed over. */
  revisionIds: readonly number[];
};

/** The operations deleting notes with their references, and the revision ids their records took. */
export type NoteDeletion = {
  ops: readonly DocumentOp[];
  revisions: readonly number[];
};

const NO_NOTE_DELETION: NoteDeletion = { ops: [], revisions: [] };

/**
 * The operations deleting the notes whose last live reference, in any
 * story, a tracked edit deleted: each note's text and paragraph marks,
 * deleted by the reference deletion's author at its date, as the reference
 * implementation records it. Rejecting the reference's deletion gives the
 * note its text back, and accepting it removes the note. A suggestion stays
 * its author's until it is accepted, so it deletes no note.
 *
 * The edit itself leaves the notes alone, so an operation touches only the
 * paragraphs it names; a compiler emitting the edit appends these.
 */
export const noteDeletionOps = ({
  before,
  after,
  stamp,
  revisionIds,
}: NoteDeletionOpsOptions): Result<NoteDeletion, DocumentOpRefusal> => {
  if (stamp.suggestionId !== undefined) return Result.ok(NO_NOTE_DELETION);
  const notes = notesLosingLiveReferences(before, after).filter(
    ({ deletion }) => carriesStamp(deletion.info, stamp),
  );
  if (notes.length === 0) return Result.ok(NO_NOTE_DELETION);
  const used = packageIdentityKeyIndex(after.package);
  const taken = new Set<string>();
  const revisionKey = (id: number) =>
    slotKey({ space: IDENTITY_SPACES.REVISION, id });
  const isFree = (id: number) =>
    !used.has(revisionKey(id)) && !taken.has(revisionKey(id));
  const ops: DocumentOp[] = [];
  for (const { story, deletion } of notes) {
    // The records take the reference deletion's id first, which the
    // reference holds: every one of them takes a fresh id.
    const deleted = deletedNoteContent(after, story, {
      ...stamp,
      id: deletion.info.id,
    });
    if (deleted === undefined) continue;
    const beforeKeys = new Set(identityKeysIn(deleted.before));
    const fresh = freshenIdentities({
      before: deleted.before,
      after: deleted.after,
      newIds: { revision: revisionIds.filter(isFree) },
      usedElsewhere: () =>
        new Set(
          [...used.counts().keys(), ...taken].filter(
            (key) => !beforeKeys.has(key),
          ),
        ),
    });
    const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
      Result.err(
        new DocumentOpRefusal({
          reason,
          message,
          opType: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        }),
      );
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
        for (const key of identityKeysIn(fresh.paragraphs)) taken.add(key);
        for (const [index, paragraph] of fresh.paragraphs.entries())
          ops.push({
            type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
            story,
            expected: [deleted.before[index] ?? paragraph],
            blocks: [paragraph],
          });
        break;
      default: {
        const unreachable: never = fresh;
        return unreachable;
      }
    }
  }
  return Result.ok({
    ops,
    revisions: revisionIds.filter(
      (id) => !used.has(revisionKey(id)) && taken.has(revisionKey(id)),
    ),
  });
};
