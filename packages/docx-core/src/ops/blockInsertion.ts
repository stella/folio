/** Paragraph insertion, isolated from the inline and table operation implementations. */
import { Result } from "better-result";

import type { Document, Paragraph } from "../model/document";
import { blockListAt, endsItsContainer, storyBody, storyParagraphs } from "./blocks";
import type { DocumentEdit } from "./edits";
import { freshenIdentities } from "./identity";
import { identityKeysIn, idKey, packageIdentityKeys } from "./ids";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import { stampInfo, WRAP_KINDS, wrapTracked } from "./review";
import { paragraphLength } from "./offsets";
import { DOCUMENT_OP_TYPES, type InsertBlocksOp } from "./types";

type BlockInsertionOptions = {
  document: Document;
  op: InsertBlocksOp;
  applyOps: ApplyOps;
};

export const insertBlocks = ({
  document,
  op,
  applyOps,
}: BlockInsertionOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  if (op.blocks.length === 0) {
    return refuse(DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST, "No paragraphs to insert.");
  }
  if (op.at.type !== "before" && op.at.type !== "after") {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "Insert before or after a paragraph.",
    );
  }
  const body = storyBody(document, op.story);
  const location = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(op.at.blockId),
  );
  if (location === undefined) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The insertion anchor does not exist.",
    );
  }
  if (op.blocks.some(({ sectionProperties }) => sectionProperties !== undefined)) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Inserting section breaks requires section operations.",
    );
  }
  // Use the following paragraph as the unchanged boundary whenever possible.
  const following = blockListAt(body.content, location.list)[location.index + 1];
  const before = op.at.type === "before" || following?.type === "paragraph";
  const anchor =
    op.at.type === "after" && following?.type === "paragraph" ? following : location.paragraph;
  if (op.revision !== undefined && !before) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Tracked insertion requires a following paragraph in the same block list.",
    );
  }
  const replacement = (blocks: readonly Paragraph[]) => ({
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: op.story,
    expected: [anchor],
    blocks: before ? [...blocks, anchor] : [anchor, ...blocks],
  });
  // Validate incoming identities and model records before creating any review records.
  const direct = applyOps(document, [replacement(op.blocks)]);
  if (direct.isErr())
    return Result.err(
      new DocumentOpRefusal({
        reason: direct.error.reason,
        message: direct.error.message,
        opType: op.type,
      }),
    );
  const insertedIds = new Set(op.blocks.map(({ paraId }) => idKey(paraId ?? "")));
  const finalMarks = storyParagraphs(storyBody(direct.value.document, op.story)).some(
    (at) =>
      insertedIds.has(idKey(at.paragraph.paraId ?? "")) &&
      at.paragraph.pPrMark !== undefined &&
      endsItsContainer(storyBody(direct.value.document, op.story), at),
  );
  if (finalMarks) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
      "Insertion would leave a tracked mark on a container-final paragraph.",
    );
  }
  const touched = {
    modified: [],
    inserted: op.blocks.flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId])),
    removed: [],
  };
  if (op.revision === undefined) return Result.ok({ ...direct.value, touched });
  if (identityKeysIn(op.blocks).length > 0) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Tracked insertion of identified review or content-control records is unsupported.",
    );
  }
  const blocks: Paragraph[] = [];
  for (const paragraph of op.blocks) {
    const wrapped = wrapTracked({
      items: paragraph.content,
      from: { offset: 0, zeroWidthBefore: 0 },
      to: { offset: paragraphLength(paragraph), zeroWidthBefore: Number.MAX_SAFE_INTEGER },
      kind: WRAP_KINDS.INSERTION,
      stamp: op.revision,
    });
    if (wrapped.kind === "refused") {
      return refuse(wrapped.reason, "The inserted content cannot be tracked.");
    }
    blocks.push({
      ...paragraph,
      content: wrapped.kind === "wrapped" ? wrapped.content : paragraph.content,
      pPrMark: { kind: "ins", info: stampInfo(op.revision) },
    });
  }
  const fresh = freshenIdentities({
    before: [],
    after: blocks,
    newIds: op.newIds ?? {},
    usedElsewhere: () => new Set(packageIdentityKeys(document.package)),
  });
  switch (fresh.kind) {
    case "needsIds":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        "Inserted paragraph review records need more revision ids.",
      );
    case "invalidId":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        "An inserted review record has an invalid revision id.",
      );
    case "fresh": {
      const tracked = applyOps(document, [replacement(fresh.paragraphs)]);
      return tracked.isErr()
        ? Result.err(
            new DocumentOpRefusal({
              reason: tracked.error.reason,
              message: tracked.error.message,
              opType: op.type,
            }),
          )
        : Result.ok({ ...tracked.value, touched });
    }
    default: {
      const unreachable: never = fresh;
      return unreachable;
    }
  }
};
