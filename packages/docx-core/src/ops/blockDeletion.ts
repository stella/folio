/** Whole-paragraph deletion within one block list, including its final paragraph. */
import { Result } from "better-result";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
  Table,
} from "../model/document";
import {
  blockListAt,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { isCommentExhausted } from "./comments";
import type { DocumentEdit } from "./edits";
import { freshenIdentities } from "./identity";
import { identityKeysIn, idKey, packageIdentityKeys } from "./ids";
import { asParagraphContent, type InlineNode } from "./leaves";
import { asMoveHalf } from "./moves";
import {
  isRangeBoundary,
  type ParagraphContentVerdict,
  paragraphHoldsContent,
  paragraphLength,
} from "./offsets";
import {
  isBlockBoundary,
  withBoundariesOutsideDeletion,
} from "./rangeBoundaries";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import {
  foldedParagraphPropertyChange,
  paragraphPropertiesOf,
  reviewFieldsOf,
  sameParagraphProperties,
  stampInfo,
  withMarkFormatting,
  withReviewFields,
  WRAP_KINDS,
  wrapTracked,
} from "./review";
import {
  type DeleteBlocksOp,
  DOCUMENT_OP_TYPES,
  type RevisionStamp,
} from "./types";

/**
 * One paragraph's content as its tracked deletion leaves it: every leaf
 * wrapped in a deletion stamped `stamp`, split around the range boundaries
 * directly in it ({@link withBoundariesOutsideDeletion}), with the boundaries
 * that stand outside. Deleting directly keeps exactly those, so accepting the
 * tracked deletion and deleting directly leave the same document.
 */
const deletedContent = (
  paragraph: Paragraph,
  stamp: RevisionStamp,
):
  | { kind: "deleted"; content: ParagraphContent[]; standing: InlineNode[] }
  | { kind: "unchanged"; standing: InlineNode[] }
  | { kind: "refused"; reason: DocumentOpRefusal["reason"] } => {
  const wrapped = wrapTracked({
    items: paragraph.content,
    from: { offset: 0, zeroWidthBefore: 0 },
    to: {
      offset: paragraphLength(paragraph),
      zeroWidthBefore: Number.MAX_SAFE_INTEGER,
    },
    kind: WRAP_KINDS.DELETION,
    stamp,
  });
  if (wrapped.kind !== "wrapped")
    return wrapped.kind === "refused"
      ? wrapped
      : { kind: "unchanged", standing: [] };
  const { nodes, standing } = withBoundariesOutsideDeletion(
    wrapped.content,
    (node) => node.type === "deletion" && node.info.id === stamp.id,
  );
  return {
    kind: "deleted",
    content: asParagraphContent(nodes),
    standing,
  };
};

/** Only what a direct deletion leaves is read off its tracked form. */
const DIRECT_DELETION_STAMP: RevisionStamp = {
  id: Number.MAX_SAFE_INTEGER,
  author: "",
  date: "",
};

/** The paragraph a table offers a mark running into it, and the tables on its path. */
export type TableRunInTarget = {
  receiver: Paragraph;
  tables: readonly Table[];
};

/**
 * The paragraph that receives `paragraph`'s content when its mark runs into
 * the following `table`: the first paragraph of its first cell, through
 * nested tables, where no pending row or cell change stands in the way.
 */
export const paragraphTableRunInTarget = (
  paragraph: Paragraph,
  table: Table,
): TableRunInTarget | undefined => {
  if (paragraph.sectionProperties !== undefined) return undefined;
  const tables: Table[] = [];
  let current = table;
  for (;;) {
    const row = current.rows.at(0);
    const cell = row?.cells.at(0);
    if (!row || !cell) return undefined;
    if (
      row.structuralChange?.type === "tableRowInsertion" ||
      row.structuralChange?.type === "tableRowDeletion" ||
      cell.structuralChange !== undefined
    ) {
      return undefined;
    }
    tables.push(current);
    const target = cell.content.at(0);
    if (target?.type === "paragraph")
      return target.sectionProperties === undefined
        ? { receiver: target, tables }
        : undefined;
    if (target?.type !== "table") return undefined;
    current = target;
  }
};

/**
 * Whether the mark of `paragraph`, which `table` follows in `blocks`, can be
 * resolved away: the block list goes on ending with a paragraph, and the
 * table has a paragraph to receive what runs in. A tracked deletion before a
 * table is admitted only where this holds, so each decision on it resolves.
 */
export const markBeforeTableResolves = ({
  blocks,
  paragraph,
  table,
}: {
  blocks: readonly BlockContent[];
  paragraph: Paragraph;
  table: Table;
}): boolean =>
  blocks.at(-1)?.type === "paragraph" &&
  paragraphTableRunInTarget(paragraph, table) !== undefined;

/**
 * Whether `block`, following a paragraph, is a block that paragraph's mark
 * neither joins nor runs into: a block content control, custom XML or a
 * preserved block. Resolving the mark away removes the paragraph only when
 * nothing is left in it ({@link paragraphMarkOutcome}).
 */
export const isOpaqueFollower = (block: BlockContent | undefined): boolean =>
  block !== undefined &&
  block.type !== "paragraph" &&
  block.type !== "table" &&
  !isBlockBoundary(block);

/** The paragraph a block offers a mark running into it. */
export type BlockRunInTarget = { receiver: Paragraph };

/**
 * The paragraph that receives `paragraph`'s content when its mark runs into
 * the following `block`, a block content control or custom XML: its first
 * paragraph, through nested ones. A control whose content is locked, or one
 * that starts with anything else, offers none.
 */
export const paragraphBlockRunInTarget = (
  paragraph: Paragraph,
  block: BlockContent,
): BlockRunInTarget | undefined => {
  if (paragraph.sectionProperties !== undefined) return undefined;
  if (block.type !== "blockSdt" && block.type !== "blockCustomXml")
    return undefined;
  if (
    block.type === "blockSdt" &&
    (block.properties.lock === "contentLocked" ||
      block.properties.lock === "sdtContentLocked")
  )
    return undefined;
  const first = block.content.at(0);
  if (first?.type !== "paragraph")
    return first === undefined
      ? undefined
      : paragraphBlockRunInTarget(paragraph, first);
  return first.sectionProperties === undefined
    ? { receiver: first }
    : undefined;
};

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
  holdsContentOnceAccepted = holdsContent,
  followingBlockCanReceiveRunIn = false,
  needsFollowingCommentHost = false,
  tableRemovalIsTrackable = true,
  removedSectionProperties = false,
  precedingSectionOwner = "unavailable",
  followingHasSectionProperties = false,
}: {
  /** `block` is a follower {@link isOpaqueFollower} admits. */
  follows: "paragraph" | "table" | "block" | "nothing";
  /** From {@link paragraphHoldsContent} or the editor's view of the same rule, never re-derived. */
  holdsContent: ParagraphContentVerdict;
  markWasAdded: boolean;
  /**
   * Read when nothing or a block follows. When nothing does: whether the
   * block BEFORE this paragraph, which becomes the container's last block
   * once this one is removed, is a paragraph; never the container's current
   * last child, which is this one. When a block does: whether the container's
   * last block is a paragraph.
   */
  containerStillEndsWithParagraph: boolean;
  followingTableCanReceiveRunIn: boolean;
  /**
   * Read when a block follows: whether the paragraph still holds content
   * once every change in it is accepted, from
   * {@link paragraphHoldsContentOnceAccepted}.
   */
  holdsContentOnceAccepted?: ParagraphContentVerdict;
  /** Read when a block follows: whether it offers a {@link paragraphBlockRunInTarget}. */
  followingBlockCanReceiveRunIn?: boolean;
  needsFollowingCommentHost?: boolean;
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
  // Only a join has a following paragraph mark. A removal whose section could
  // only move forward keeps the paragraph instead, so its section survives.
  const removalWouldDropSection = sectionDestination === "following";
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
      if (!holdsContent) {
        if (removalWouldDropSection)
          return { kind: "keep", sectionDestination: "none" };
        return needsFollowingCommentHost && followingTableCanReceiveRunIn
          ? { kind: "run-in", sectionDestination }
          : { kind: "remove", sectionDestination };
      }
      return followingTableCanReceiveRunIn
        ? { kind: "run-in", sectionDestination }
        : { kind: "keep", sectionDestination };
    case "block":
      // An accepted deletion's paragraph goes once nothing is left in it,
      // and what accepting will take from it still runs into the block, so
      // resolving its content before or after its mark ends the same. What
      // stays once accepted keeps the paragraph, without its mark.
      if (removalWouldDropSection)
        return { kind: "keep", sectionDestination: "none" };
      if (markWasAdded) return { kind: "keep", sectionDestination };
      if (!holdsContent)
        return containerStillEndsWithParagraph
          ? { kind: "remove", sectionDestination }
          : { kind: "keep", sectionDestination };
      return !holdsContentOnceAccepted && followingBlockCanReceiveRunIn
        ? { kind: "run-in", sectionDestination }
        : { kind: "keep", sectionDestination };
    case "nothing":
      if (removalWouldDropSection)
        return { kind: "keep", sectionDestination: "none" };
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
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
      "No paragraphs to delete.",
    );
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
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
      "No paragraphs to delete.",
    );
  }
  if (
    !found.every(
      (location, index) =>
        sameBlockList(location.list, first.list) &&
        location.index === first.index + index,
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
  // Tracked, a paragraph a table or another block follows loses its words
  // and its own mark, as the legacy editor records it; accepting removes
  // it, which needs what resolving that mark needs.
  const trackedBeforeTable =
    op.revision !== undefined &&
    following?.type === "table" &&
    markBeforeTableResolves({
      blocks: list,
      paragraph: last.paragraph,
      table: following,
    });
  const trackedBeforeBlock =
    op.revision !== undefined &&
    isOpaqueFollower(following) &&
    list.at(-1)?.type === "paragraph";
  if (
    following?.type !== "paragraph" &&
    !terminal &&
    !trackedBeforeTable &&
    !trackedBeforeBlock
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Deletion requires a following paragraph or a container-final paragraph.",
    );
  }
  const preceding = terminal ? list[first.index - 1] : undefined;
  const previous = preceding?.type === "paragraph" ? preceding : undefined;
  const selected = found.map(({ paragraph }) => paragraph);
  const affected = previous === undefined ? selected : [previous, ...selected];
  if (
    affected.some(({ sectionProperties }) => sectionProperties !== undefined)
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Deleting section boundaries requires section operations.",
    );
  }
  const replacement = (blocks: readonly Paragraph[]) => ({
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: op.story,
    expected:
      following?.type === "paragraph" ? [...selected, following] : affected,
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
  // Deleting a paragraph takes its content, never a position a user placed
  // in it: the boundaries its tracked deletion leaves outside stay, where the
  // deleted paragraphs joined what follows.
  const kept: ParagraphContent[] = [];
  for (const paragraph of selected) {
    const form = deletedContent(
      paragraph,
      op.revision ?? DIRECT_DELETION_STAMP,
    );
    if (form.kind === "refused" && op.revision !== undefined)
      return refuse(
        form.reason,
        "The selected paragraph content cannot be tracked.",
      );
    // Content no tracked deletion can hold still goes directly; the
    // boundaries standing in the paragraph itself stay.
    kept.push(
      ...(form.kind === "refused"
        ? paragraph.content.filter(isRangeBoundary)
        : asParagraphContent(form.standing)),
    );
  }
  const survivor = withReviewFields(
    { ...last.paragraph, content: [...(previous?.content ?? []), ...kept] },
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
  if (outcome.kind === "join" && following?.type === "paragraph")
    directBlocks = [
      kept.length === 0
        ? following
        : { ...following, content: [...kept, ...following.content] },
    ];
  // A deletion is refused before a table and keeps the paragraph that ends
  // its container, so no outcome here removes the boundaries it kept: a
  // removal, were one decided, keeps the survivor holding them.
  if (outcome.kind === "keep" || outcome.kind === "remove")
    directBlocks = [survivor];
  const direct = applyOps(document, [replacement(directBlocks)]);
  if (direct.isErr()) return refuse(direct.error.reason, direct.error.message);
  if (op.revision === undefined) {
    let edit = direct.value;
    let retiredTouched: readonly string[] = [];
    const exhaustedCommentIds = (
      document.package.document.comments ?? []
    ).flatMap(({ id, parentId }) =>
      parentId === undefined &&
      !isCommentExhausted({ type: "document", document, id }) &&
      isCommentExhausted({
        type: "document",
        document: direct.value.document,
        id,
      })
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
      if (retired.isErr())
        return refuse(retired.error.reason, retired.error.message);
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
          ...new Set([
            ...retiredTouched,
            ...(terminal ? [last.paragraph.paraId ?? ""] : []),
          ]),
        ],
        inserted: [],
        removed: affected
          .filter((paragraph) => !terminal || paragraph !== last.paragraph)
          .map(({ paraId }) => paraId ?? ""),
      },
    });
  }
  const revision = op.revision;
  // The paragraphs whose own mark the tracked deletion records deleted.
  const losesMark = (paragraph: Paragraph) =>
    !terminal || paragraph !== last.paragraph;
  // A mark the same author's pending insertion or move added goes outright,
  // joining its paragraph to the next one, as deleting over one's own
  // insertion removes it; no deletion of an addition is recorded. Before a
  // table, the last paragraph has none to join and keeps the refusal.
  const joinsNext = new Set(
    affected.filter(
      (paragraph) =>
        losesMark(paragraph) &&
        (paragraph !== last.paragraph || following?.type === "paragraph") &&
        (paragraph.pPrMark?.kind === "ins" ||
          paragraph.pPrMark?.kind === "moveTo") &&
        paragraph.pPrMark.info.author === revision.author,
    ),
  );
  if (
    affected.some(
      (paragraph) =>
        paragraph.pPrMark !== undefined && !joinsNext.has(paragraph),
    )
  ) {
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
      const form = deletedContent(paragraph, op.revision);
      if (form.kind === "refused")
        return refuse(
          form.reason,
          "The selected paragraph content cannot be tracked.",
        );
      if (form.kind === "deleted")
        next = {
          ...paragraph,
          content:
            op.move === true
              ? asMoveHalf(form.content, op.revision.id, "from")
              : form.content,
        };
    }
    if (joinsNext.has(paragraph)) {
      const { pPrMark: _removed, ...unmarked } = next;
      next = unmarked;
    } else if (losesMark(paragraph)) {
      next = {
        ...next,
        // A move's source marks its paragraph moved from, not deleted.
        pPrMark: {
          kind: op.move === true ? "moveFrom" : "del",
          info: stampInfo(op.revision),
        },
      };
    } else if (!sameParagraphProperties(next.formatting, survivor.formatting)) {
      const review = reviewFieldsOf(next);
      delete review.formatting;
      if (survivor.formatting !== undefined)
        review.formatting = survivor.formatting;
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
      new Set(
        packageIdentityKeys(document.package).filter(
          (key) => !beforeKeys.has(key),
        ),
      ),
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
      const marked =
        following?.type === "paragraph"
          ? [...fresh.paragraphs, following]
          : fresh.paragraphs;
      const original =
        following?.type === "paragraph" ? [...affected, following] : affected;
      const changed =
        following?.type === "paragraph" ? [...tracked, following] : tracked;
      // Each paragraph whose mark went joins the one after it, which keeps
      // its own identity and properties.
      const blocks: Paragraph[] = [];
      const removed: string[] = [];
      const modified = new Set<string>();
      let carried: ParagraphContent[] | undefined;
      for (const [index, paragraph] of marked.entries()) {
        const source = original[index];
        if (source !== undefined && joinsNext.has(source)) {
          carried = [...(carried ?? []), ...paragraph.content];
          removed.push(paragraph.paraId ?? "");
          continue;
        }
        const joined =
          carried === undefined
            ? paragraph
            : { ...paragraph, content: [...carried, ...paragraph.content] };
        if (carried !== undefined || changed[index] !== source)
          modified.add(joined.paraId ?? "");
        carried = undefined;
        blocks.push(joined);
      }
      const result = applyOps(document, [replacement(blocks)]);
      if (result.isErr())
        return refuse(result.error.reason, result.error.message);
      return Result.ok({
        ...result.value,
        touched: { modified: [...modified], inserted: [], removed },
      });
    }
    default: {
      const unreachable: never = fresh;
      return unreachable;
    }
  }
};
