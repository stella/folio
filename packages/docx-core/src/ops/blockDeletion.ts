/** Whole-paragraph deletion within one block list, including its final paragraph. */
import { Result } from "better-result";

import type { Document, Paragraph } from "../model/document";
import { blockListAt, sameBlockList, storyBody, storyParagraphs } from "./blocks";
import { isCommentExhausted } from "./comments";
import type { DocumentEdit } from "./edits";
import { freshenIdentities } from "./identity";
import { identityKeysIn, idKey, packageIdentityKeys } from "./ids";
import { paragraphHoldsContent, paragraphLength } from "./offsets";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import {
  paragraphPropertiesOf,
  foldedParagraphPropertyChange,
  reviewFieldsOf,
  sameParagraphProperties,
  stampInfo,
  withMarkFormatting,
  withReviewFields,
  WRAP_KINDS,
  wrapTracked,
} from "./review";
import { DOCUMENT_OP_TYPES, type DeleteBlocksOp } from "./types";

type BlockDeletionOptions = {
  document: Document;
  op: DeleteBlocksOp;
  applyOps: ApplyOps;
};

/** The single structural outcome of resolving a paragraph mark. */
export const paragraphMarkOutcome = ({
  follows,
  holdsContent,
  markWasAdded,
  containerStillEndsWithParagraph,
  followingTableCanReceiveRunIn,
  tableRemovalIsTrackable = true,
  removedSectionProperties = false,
  precedingSectionOwner = "unavailable",
  followingHasSectionProperties = false,
}: {
  follows: "paragraph" | "table" | "nothing";
  holdsContent: boolean;
  markWasAdded: boolean;
  /**
   * Read only when nothing follows: whether the block BEFORE this paragraph,
   * which becomes the container's last block once this one is removed, is a
   * paragraph. Never the container's current last child, which is this one.
   */
  containerStillEndsWithParagraph: boolean;
  followingTableCanReceiveRunIn: boolean;
  tableRemovalIsTrackable?: boolean;
  removedSectionProperties?: boolean;
  precedingSectionOwner?: "unavailable" | "available" | "same" | "occupied";
  followingHasSectionProperties?: boolean;
}):
  | ({ kind: "join" } & SectionTransfer)
  | ({ kind: "remove" } & SectionTransfer)
  | ({ kind: "run-in" } & SectionTransfer)
  | ({ kind: "keep" } & SectionTransfer)
  | ({ kind: "refuse"; reason: "untrackable" } & SectionTransfer) => {
  let sectionDestination: SectionTransfer["sectionDestination"] = "none";
  if (removedSectionProperties && precedingSectionOwner === "available") {
    sectionDestination = "preceding";
  } else if (
    removedSectionProperties &&
    precedingSectionOwner !== "same" &&
    !followingHasSectionProperties
  ) {
    sectionDestination = "following";
  }
  switch (follows) {
    case "paragraph":
      return { kind: "join", sectionDestination };
    case "table":
      if (!tableRemovalIsTrackable)
        return {
          kind: "refuse",
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          sectionDestination,
        };
      if (!holdsContent) return { kind: "remove", sectionDestination };
      return followingTableCanReceiveRunIn
        ? { kind: "run-in", sectionDestination }
        : { kind: "keep", sectionDestination };
    case "nothing":
      return !holdsContent && markWasAdded && containerStillEndsWithParagraph
        ? { kind: "remove", sectionDestination }
        : { kind: "keep", sectionDestination };
  }
};

type SectionTransfer = {
  sectionDestination: "none" | "preceding" | "following";
};

export const deleteBlocks = ({
  document,
  op,
  applyOps,
}: BlockDeletionOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  if (op.blockIds.length === 0) {
    return refuse(DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST, "No paragraphs to delete.");
  }
  const body = storyBody(document, op.story);
  const paragraphs = storyParagraphs(body);
  const found = [];
  for (const id of op.blockIds) {
    const location = paragraphs.find(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(id),
    );
    if (location === undefined) {
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        "A selected paragraph does not exist.",
      );
    }
    found.push(location);
  }
  const first = found.at(0);
  const last = found.at(-1);
  if (first === undefined || last === undefined) {
    return refuse(DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST, "No paragraphs to delete.");
  }
  if (
    !found.every(
      (location, index) =>
        sameBlockList(location.list, first.list) && location.index === first.index + index,
    )
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
      "Selected paragraphs must be adjacent in one block list.",
    );
  }
  const list = blockListAt(body.content, first.list);
  const following = list[last.index + 1];
  const terminal = last.index === list.length - 1;
  if (following?.type !== "paragraph" && !terminal) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Deletion requires a following paragraph or a container-final paragraph.",
    );
  }
  const preceding = terminal ? list[first.index - 1] : undefined;
  const previous = preceding?.type === "paragraph" ? preceding : undefined;
  const selected = found.map(({ paragraph }) => paragraph);
  const affected = previous === undefined ? selected : [previous, ...selected];
  if (affected.some(({ sectionProperties }) => sectionProperties !== undefined)) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Deleting section boundaries requires section operations.",
    );
  }
  const replacement = (blocks: readonly Paragraph[]) => ({
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: op.story,
    expected: following?.type === "paragraph" ? [...selected, following] : affected,
    blocks,
  });
  const finalReview = reviewFieldsOf(last.paragraph);
  if (previous !== undefined && paragraphLength(previous) > 0) {
    const formatting = withMarkFormatting(
      paragraphPropertiesOf(previous.formatting),
      last.paragraph.formatting,
    );
    delete finalReview.formatting;
    if (formatting !== undefined) finalReview.formatting = formatting;
  }
  const survivor = withReviewFields(
    { ...last.paragraph, content: previous?.content ?? [] },
    finalReview,
  );
  let follows: "paragraph" | "table" | "nothing" = "nothing";
  if (following?.type === "paragraph") follows = "paragraph";
  if (following?.type === "table") follows = "table";
  const outcome = paragraphMarkOutcome({
    follows,
    holdsContent: paragraphHoldsContent(survivor),
    markWasAdded: false,
    containerStillEndsWithParagraph: previous !== undefined,
    followingTableCanReceiveRunIn: false,
  });
  let directBlocks: Paragraph[] = [];
  if (outcome.kind === "join" && following?.type === "paragraph") directBlocks = [following];
  if (outcome.kind === "keep") directBlocks = [survivor];
  const direct = applyOps(document, [replacement(directBlocks)]);
  if (direct.isErr()) return refuse(direct.error.reason, direct.error.message);
  if (op.revision === undefined) {
    let edit = direct.value;
    let retiredTouched: readonly string[] = [];
    const exhaustedCommentIds = (document.package.document.comments ?? []).flatMap(
      ({ id, parentId }) =>
        parentId === undefined &&
        !isCommentExhausted({ type: "document", document, id }) &&
        isCommentExhausted({ type: "document", document: direct.value.document, id })
          ? [id]
          : [],
    );
    if (exhaustedCommentIds.length > 0) {
      const retired = applyOps(
        direct.value.document,
        exhaustedCommentIds.map((id) => ({
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id,
          scope: "thread" as const,
        })),
      );
      if (retired.isErr()) return refuse(retired.error.reason, retired.error.message);
      retiredTouched = retired.value.touched.modified;
      edit = {
        document: retired.value.document,
        inverse: [...retired.value.inverse, ...direct.value.inverse],
        touched: retired.value.touched,
      };
    }
    // Anchoring the inverse must not report an unchanged boundary as modified.
    return Result.ok({
      ...edit,
      touched: {
        modified: [
          ...new Set([...retiredTouched, ...(terminal ? [last.paragraph.paraId ?? ""] : [])]),
        ],
        inserted: [],
        removed: affected
          .filter((paragraph) => !terminal || paragraph !== last.paragraph)
          .map(({ paraId }) => paraId ?? ""),
      },
    });
  }
  if (affected.some(({ pPrMark }) => pPrMark !== undefined)) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      "Deletion would overwrite a paragraph review record.",
    );
  }
  if (identityKeysIn(selected.map(({ content }) => ({ content }))).length > 0) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Tracked paragraph deletion of identified review or content-control content is unsupported.",
    );
  }
  const selectedSet = new Set(selected);
  const tracked: Paragraph[] = [];
  for (const paragraph of affected) {
    let next = paragraph;
    if (selectedSet.has(paragraph)) {
      const wrapped = wrapTracked({
        items: paragraph.content,
        from: { offset: 0, zeroWidthBefore: 0 },
        to: { offset: paragraphLength(paragraph), zeroWidthBefore: Number.MAX_SAFE_INTEGER },
        kind: WRAP_KINDS.DELETION,
        stamp: op.revision,
      });
      if (wrapped.kind === "refused")
        return refuse(wrapped.reason, "The selected paragraph content cannot be tracked.");
      if (wrapped.kind === "wrapped") next = { ...paragraph, content: wrapped.content };
    }
    if (!terminal || paragraph !== last.paragraph) {
      next = { ...next, pPrMark: { kind: "del", info: stampInfo(op.revision) } };
    } else if (!sameParagraphProperties(next.formatting, survivor.formatting)) {
      const review = reviewFieldsOf(next);
      delete review.formatting;
      if (survivor.formatting !== undefined) review.formatting = survivor.formatting;
      review.propertyChanges = [
        foldedParagraphPropertyChange({
          paragraph,
          formatting: survivor.formatting,
          stamp: op.revision,
        }),
      ];
      next = withReviewFields(next, review);
    }
    tracked.push(next);
  }
  const beforeKeys = new Set(identityKeysIn(affected));
  const fresh = freshenIdentities({
    before: affected,
    after: tracked,
    newIds: op.newIds ?? {},
    usedElsewhere: () =>
      new Set(packageIdentityKeys(document.package).filter((key) => !beforeKeys.has(key))),
  });
  switch (fresh.kind) {
    case "needsIds":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        "Paragraph deletion needs more revision ids.",
      );
    case "invalidId":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        "Paragraph deletion received an invalid revision id.",
      );
    case "fresh": {
      const blocks =
        following?.type === "paragraph" ? [...fresh.paragraphs, following] : fresh.paragraphs;
      const result = applyOps(document, [replacement(blocks)]);
      if (result.isErr()) return refuse(result.error.reason, result.error.message);
      return Result.ok({
        ...result.value,
        touched: {
          modified: tracked
            .filter((paragraph, index) => paragraph !== affected[index])
            .map(({ paraId }) => paraId ?? ""),
          inserted: [],
          removed: [],
        },
      });
    }
    default: {
      const unreachable: never = fresh;
      return unreachable;
    }
  }
};
