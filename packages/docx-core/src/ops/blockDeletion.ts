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
import { structurallyEqual } from "./equality";
import { freshenIdentities } from "./identity";
import { identityKeysIn, idKey, packageIdentityKeys } from "./ids";
import {
  asParagraphContent,
  childNodes,
  type InlineNode,
  isCommentAnchor,
} from "./leaves";
import { asMoveHalf } from "./moves";
import {
  isRangeBoundary,
  type ParagraphContentVerdict,
  paragraphHoldsContent,
  paragraphHoldsContentOnceAccepted,
  paragraphLength,
} from "./offsets";
import {
  isBlockBoundary,
  isBoundary,
  withBoundariesOutsideDeletion,
} from "./rangeBoundaries";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import {
  FOLD_INTO_OTHER_SUGGESTION_MESSAGE,
  foldedParagraphPropertyChange,
  foldsIntoOtherSuggestion,
  isRemovedRevisionNode,
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
  isOwnRecord,
  recordedOnlyBy,
  type RevisionOwner,
} from "./revisionOwner";
import {
  type DeleteBlocksOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type ParagraphAbsorption,
  type RevisionStamp,
  SECTION_BOUNDARY_POLICIES,
} from "./types";

/**
 * One paragraph's content as its tracked deletion leaves it: every leaf
 * wrapped in a deletion stamped `stamp`, split around the range boundaries
 * directly in it ({@link withBoundariesOutsideDeletion}), with the boundaries
 * that stand outside, the bare comment anchors the deletion leaves standing
 * among them, in document order. Deleting directly keeps exactly those, so
 * accepting the tracked deletion and deleting directly leave the same
 * document.
 */
export const deletedContent = (
  paragraph: Paragraph,
  stamp: RevisionStamp,
  rangedComments: ReadonlySet<number>,
):
  | { kind: "deleted"; content: ParagraphContent[]; standing: InlineNode[] }
  | { kind: "unchanged"; standing: InlineNode[] }
  | { kind: "refused"; reason: DocumentOpRefusal["reason"] } => {
  // A comment's reference stands with its range's end, which the deletion
  // leaves standing wherever that end is; a point comment's reference is its
  // only anchor, which a deletion takes.
  if (!liveReferencesAreRanged(paragraph.content, rangedComments))
    return { kind: "refused", reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE };
  const wrapped = wrapTracked({
    items: paragraph.content,
    from: { offset: 0, zeroWidthBefore: 0 },
    to: {
      offset: paragraphLength(paragraph),
      zeroWidthBefore: Number.MAX_SAFE_INTEGER,
    },
    kind: WRAP_KINDS.DELETION,
    stamp,
    commentAnchors: "stand",
  });
  if (wrapped.kind === "refused") return wrapped;
  if (wrapped.kind === "unchanged")
    return {
      kind: "unchanged",
      standing: standingIn(paragraph.content, new Set(), () => false),
    };
  const created = (node: InlineNode) =>
    node.type === "deletion" && node.info.id === stamp.id;
  const { nodes, standing } = withBoundariesOutsideDeletion(
    wrapped.content,
    created,
  );
  return {
    kind: "deleted",
    content: asParagraphContent(nodes),
    standing: standingIn(nodes, new Set(standing), created),
  };
};

/** Whether accepting every change takes `paragraph`'s own mark. */
const markGoes = ({ pPrMark }: Paragraph): boolean =>
  pPrMark?.kind === "del" ||
  pPrMark?.kind === "moveFrom" ||
  pPrMark?.deletion !== undefined;

/**
 * Whether `owner`'s pending insertion added `paragraph` whole: its mark
 * and every word, with no one else's suggestion in it and no range boundary
 * standing in it, which removing the paragraph would lose.
 */
const addedWholeBy = (paragraph: Paragraph, owner: RevisionOwner): boolean =>
  paragraph.pPrMark?.kind === "ins" &&
  paragraph.pPrMark.deletion === undefined &&
  isOwnRecord(paragraph.pPrMark.info, owner) &&
  paragraph.content.every(
    (item) => item.type === "insertion" && isOwnRecord(item.info, owner),
  ) &&
  !holdsRangeBoundary(paragraph.content) &&
  recordedOnlyBy(paragraph, owner);

/** Whether a range boundary stands anywhere among `nodes`. */
const holdsRangeBoundary = (nodes: readonly InlineNode[]): boolean =>
  nodes.some(
    (node) => isBoundary(node) || holdsRangeBoundary(childNodes(node) ?? []),
  );

/**
 * Every comment reference in `nodes` outside a removal already pending
 * belongs to a comment with a range (`rangedComments`); one inside such a
 * removal goes with it either way.
 */
const liveReferencesAreRanged = (
  nodes: readonly InlineNode[],
  rangedComments: ReadonlySet<number>,
): boolean =>
  nodes.every((node) => {
    if (isRemovedRevisionNode(node)) return true;
    if (node.type === "commentReference") return rangedComments.has(node.id);
    return liveReferencesAreRanged(childNodes(node) ?? [], rangedComments);
  });

/** The ids of the comments whose range ends somewhere in `value`. */
export const rangeEndIdsIn = (
  value: unknown,
  ids = new Set<number>(),
): Set<number> => {
  if (Array.isArray(value)) {
    for (const item of value) rangeEndIdsIn(item, ids);
    return ids;
  }
  if (typeof value !== "object" || value === null) return ids;
  const id: unknown = Reflect.get(value, "id");
  if (
    Reflect.get(value, "type") === "commentRangeEnd" &&
    typeof id === "number"
  )
    ids.add(id);
  for (const item of Object.values(value)) rangeEndIdsIn(item, ids);
  return ids;
};

/**
 * The boundaries of `nodes` a deletion of their paragraph leaves standing,
 * in document order: those split out of its records (`split`) and the bare
 * comment range anchors it never took in. What a deletion stamped by the same
 * revision (`created`) or an earlier one holds goes with it.
 */
const standingIn = (
  nodes: readonly InlineNode[],
  split: ReadonlySet<InlineNode>,
  created: (node: InlineNode) => boolean,
): InlineNode[] =>
  nodes.flatMap((node): InlineNode[] => {
    if (split.has(node) || isCommentAnchor(node)) return [node];
    if (created(node) || isRemovedRevisionNode(node)) return [];
    return standingIn(childNodes(node) ?? [], split, created);
  });

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

type TableRunInBoundary =
  | { type: "receiver"; target: TableRunInTarget }
  | { type: "blocked"; reason: "section" | "review" | "structure" };

/** The first-cell path, including why it cannot receive paragraph content. */
export const paragraphTableRunInBoundary = (
  paragraph: Paragraph,
  table: Table,
): TableRunInBoundary => {
  if (paragraph.sectionProperties !== undefined)
    return { type: "blocked", reason: "section" };
  const tables: Table[] = [];
  let current = table;
  let review: "clear" | "pending" = "clear";
  for (;;) {
    const row = current.rows.at(0),
      cell = row?.cells.at(0);
    if (row === undefined || cell === undefined)
      return { type: "blocked", reason: "structure" };
    if (
      row.structuralChange?.type === "tableRowInsertion" ||
      row.structuralChange?.type === "tableRowDeletion" ||
      cell.structuralChange !== undefined
    )
      review = "pending";
    tables.push(current);
    const target = cell.content.at(0);
    if (target?.type === "paragraph") {
      if (target.sectionProperties !== undefined)
        return { type: "blocked", reason: "section" };
      if (review === "pending") return { type: "blocked", reason: "review" };
      return { type: "receiver", target: { receiver: target, tables } };
    }
    if (target?.type !== "table")
      return { type: "blocked", reason: "structure" };
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
  boundary,
}: {
  blocks: readonly BlockContent[];
  boundary: TableRunInBoundary;
}): boolean =>
  blocks.at(-1)?.type === "paragraph" && boundary.type === "receiver";

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
  emptyInsertedTableBoundary = "unavailable",
  precededByBlock = false,
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
  /** Only a pending-review boundary proved free of source section/structure blockers. */
  emptyInsertedTableBoundary?: "unavailable" | "pendingReview";
  /** Whether a block stays before the paragraph to anchor what its removal leaves. */
  precededByBlock?: boolean;
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
      // An accepted deletion that leaves the paragraph empty is no run-in
      // source either: it goes without crossing the table, wherever the
      // story ends, as it does before a table that can receive words.
      if (
        !tableRemovalIsTrackable &&
        precededByBlock &&
        !markWasAdded &&
        !holdsContent &&
        !needsFollowingCommentHost
      )
        return removalWouldDropSection
          ? { kind: "keep", sectionDestination: "none" }
          : { kind: "remove", sectionDestination };
      // A rejected insertion owns an empty paragraph, not a run-in source.
      // Its mark can go without crossing the table when no comment host is
      // retained; existing marks and content still require a trackable boundary.
      if (
        !tableRemovalIsTrackable &&
        (emptyInsertedTableBoundary !== "pendingReview" ||
          !markWasAdded ||
          holdsContent ||
          needsFollowingCommentHost)
      )
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
      // A resolved-away empty mark may go only while the container still
      // ends with a paragraph. Rejecting an added mark keeps standing words;
      // an accepted deletion may run its still-pending words into the block.
      if (removalWouldDropSection)
        return { kind: "keep", sectionDestination: "none" };
      if (!holdsContent)
        return containerStillEndsWithParagraph
          ? { kind: "remove", sectionDestination }
          : { kind: "keep", sectionDestination };
      if (markWasAdded) return { kind: "keep", sectionDestination };
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

type DirectRemovalReceiverOptions = {
  before: BlockContent | undefined;
  following: BlockContent | undefined;
  kept: readonly ParagraphContent[];
};

/**
 * The paragraph that takes the boundaries of paragraphs a direct deletion
 * removes before a block, with its content once it holds them: the one
 * before them, at its end, else the block's first paragraph, at its start.
 */
const directRemovalReceiver = ({
  before,
  following,
  kept,
}: DirectRemovalReceiverOptions):
  { receiver: Paragraph; content: ParagraphContent[] } | undefined => {
  if (before?.type === "paragraph")
    return { receiver: before, content: [...before.content, ...kept] };
  if (following === undefined) return undefined;
  const carrier: Paragraph = { type: "paragraph", content: [...kept] };
  let receiver: Paragraph | undefined;
  if (following.type === "table") {
    const boundary = paragraphTableRunInBoundary(carrier, following);
    if (boundary.type === "receiver") receiver = boundary.target.receiver;
  } else receiver = paragraphBlockRunInTarget(carrier, following)?.receiver;
  return receiver === undefined
    ? undefined
    : { receiver, content: [...kept, ...receiver.content] };
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
  const rangedComments = rangeEndIdsIn(body.content);
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
  // Paragraphs the same author's pending insertion added whole, with no one
  // else's suggestion in them, go outright before a table, as deleting over
  // one's own insertion removes it; no deletion of an addition is recorded.
  const owner = op.revision;
  if (
    owner !== undefined &&
    following?.type === "table" &&
    list.at(-1)?.type === "paragraph" &&
    found.every(({ paragraph }) => addedWholeBy(paragraph, owner))
  ) {
    const removed = applyOps(document, [
      {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: op.story,
        expected: found.map(({ paragraph }) => paragraph),
        blocks: [],
      },
    ]);
    if (removed.isErr())
      return refuse(removed.error.reason, removed.error.message);
    return Result.ok({
      document: removed.value.document,
      inverse: removed.value.inverse,
      touched: {
        modified: [],
        inserted: [],
        removed: found.map(({ paragraph }) => paragraph.paraId ?? ""),
      },
    });
  }
  // Tracked, a paragraph a table or another block follows loses its words
  // and its own mark, as the default editor records it; accepting removes
  // it, which needs what resolving that mark needs.
  const tableBoundary =
    following?.type === "table"
      ? paragraphTableRunInBoundary(last.paragraph, following)
      : undefined;
  // Before a table whose first row is itself pending review, the words have
  // nowhere to run in, but the deletion takes them all: accepting leaves the
  // paragraph empty, and an empty paragraph goes without crossing the table,
  // first in its container or after another block, as resolving its mark
  // already does.
  const emptiedBeforeTableUnderReview =
    tableBoundary?.type === "blocked" &&
    tableBoundary.reason === "review" &&
    list.at(-1)?.type === "paragraph";
  const trackedBeforeTable =
    op.revision !== undefined &&
    tableBoundary !== undefined &&
    (markBeforeTableResolves({ blocks: list, boundary: tableBoundary }) ||
      emptiedBeforeTableUnderReview);
  const trackedBeforeBlock =
    op.revision !== undefined &&
    isOpaqueFollower(following) &&
    list.at(-1)?.type === "paragraph";
  // Directly, paragraphs a table or another block follows, with no table
  // before them, go whole, as accepting their tracked deletion removes them:
  // the container still ends with a paragraph, and no two tables meet.
  const directBeforeBlock =
    op.revision === undefined &&
    following !== undefined &&
    following.type !== "paragraph" &&
    list.at(-1)?.type === "paragraph" &&
    (first.index === 0 || list[first.index - 1]?.type === "paragraph") &&
    found.every(({ paragraph }) => paragraph.sectionProperties === undefined);
  if (
    following?.type !== "paragraph" &&
    !terminal &&
    !trackedBeforeTable &&
    !trackedBeforeBlock &&
    !directBeforeBlock
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Deletion requires a following paragraph or a container-final paragraph.",
    );
  }
  // Tracked, paragraphs a pending deletion already takes whole, words and
  // mark, stand between the container's last paragraph and the break its
  // deletion retires: that is the break before them, as accepting them would
  // leave it.
  let previousIndex = first.index - 1;
  if (terminal && op.revision !== undefined) {
    let candidate = previousIndex;
    for (;;) {
      const block = list[candidate];
      if (
        block?.type !== "paragraph" ||
        block.pPrMark?.kind !== "del" ||
        block.sectionProperties !== undefined ||
        paragraphHoldsContentOnceAccepted(block)
      )
        break;
      candidate -= 1;
    }
    if (candidate !== previousIndex && list[candidate]?.type === "paragraph")
      previousIndex = candidate;
  }
  const preceding = terminal ? list[previousIndex] : undefined;
  const previous = preceding?.type === "paragraph" ? preceding : undefined;
  const between = list
    .slice(previousIndex + 1, first.index)
    .filter((block): block is Paragraph => block.type === "paragraph");
  const selected = found.map(({ paragraph }) => paragraph);
  const affected =
    previous === undefined ? selected : [previous, ...between, ...selected];
  // A section break is a property of the paragraph mark that ends the
  // section. A direct deletion that takes that mark takes the break: the
  // sections on either side become one under the properties of the section
  // after it, as accepting a tracked deletion of the mark leaves them. A
  // tracked deletion records the mark deleted with its break on it, so
  // rejecting it keeps the section as it was.
  const expected =
    following?.type === "paragraph" ? [...selected, following] : affected;
  const sectionBoundaries = expected.some(
    ({ sectionProperties }) => sectionProperties !== undefined,
  )
    ? { sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE }
    : {};
  const replacement = (blocks: readonly Paragraph[]) => ({
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: op.story,
    expected,
    blocks,
    ...sectionBoundaries,
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
      rangedComments,
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
  const directModified: string[] = [];
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
  const directOps: DocumentOp[] = [];
  // Removed before a block, they go whole, as accepting their tracked
  // deletion removes them: the boundaries they kept go to the end of the
  // paragraph before them, else run into the block's first paragraph.
  if (directBeforeBlock) {
    directBlocks = [];
    if (kept.length > 0) {
      const received = directRemovalReceiver({
        before: list[first.index - 1],
        following,
        kept,
      });
      if (received === undefined)
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          "Deletion requires a following paragraph or a container-final paragraph.",
        );
      directOps.push({
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: op.story,
        expected: [received.receiver],
        blocks: [{ ...received.receiver, content: received.content }],
      });
      directModified.push(received.receiver.paraId ?? "");
    }
  }
  directOps.push(replacement(directBlocks));
  const direct = applyOps(document, directOps);
  if (direct.isErr()) return refuse(direct.error.reason, direct.error.message);
  // A paragraph the direct deletion writes back is modified exactly when it
  // is not the block it replaces: the survivor, or the paragraph after the
  // deletion once it holds the boundaries the deleted paragraphs kept.
  const replaced = new Map(
    replacement([]).expected.map((block) => [idKey(block.paraId ?? ""), block]),
  );
  directModified.push(
    ...directBlocks
      .filter((block) => replaced.get(idKey(block.paraId ?? "")) !== block)
      .map(({ paraId }) => paraId ?? ""),
  );
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
        modified: [...new Set([...retiredTouched, ...directModified])],
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
  // table, the last paragraph has none to join and keeps the refusal. A mark
  // that ends a section is not taken back here: where its break goes is what
  // rejecting the insertion decides, so the deletion is recorded beside it
  // and that resolution moves the break.
  const joinsNext = new Set(
    affected.filter(
      (paragraph) =>
        losesMark(paragraph) &&
        paragraph.sectionProperties === undefined &&
        (paragraph !== last.paragraph || following?.type === "paragraph") &&
        (paragraph.pPrMark?.kind === "ins" ||
          paragraph.pPrMark?.kind === "moveTo") &&
        isOwnRecord(paragraph.pPrMark.info, revision),
    ),
  );
  // A mark a pending deletion already records deleted stays as it is: the
  // break goes either way, under the deletion that took it first.
  // Not where that deletion took the mark for the container's last
  // paragraph, whose words it deleted and to which it gave this paragraph's
  // properties: deleting this paragraph too leaves those properties on a
  // paragraph nothing of this one reaches.
  const finalTakeover =
    following?.type === "paragraph" &&
    last.index + 1 === list.length - 1 &&
    paragraphHoldsContent(following) &&
    !paragraphHoldsContentOnceAccepted(following) &&
    (following.propertyChanges?.length ?? 0) > 0 &&
    sameParagraphProperties(
      paragraphPropertiesOf(following.formatting),
      paragraphPropertiesOf(last.paragraph.formatting),
    );
  // So does an added mark whose later deletion already took it.
  const keepsDeletedMark = (paragraph: Paragraph) =>
    op.move !== true &&
    (paragraph.pPrMark?.kind === "del" ||
      paragraph.pPrMark?.deletion !== undefined) &&
    !(finalTakeover && paragraph === last.paragraph);
  // Another owner's added mark the deletion takes keeps that insertion and
  // records the deletion beside it, as a later deletion of an insertion does;
  // so does one's own that ends a section, and one's own on the last
  // paragraph before a table whose rows are all pending insertion: it has no
  // paragraph to join, and accepting both records removes it without running
  // into the table. Not where the paragraph before it loses its own mark:
  // accepting would run that one into the table instead.
  const before = list[first.index - 1];
  const beforeInsertedTable =
    !(before?.type === "paragraph" && markGoes(before)) &&
    following?.type === "table" &&
    following.rows.length > 0 &&
    following.rows.every(
      (row) => row.structuralChange?.type === "tableRowInsertion",
    );
  const deletesAddedMark = (paragraph: Paragraph) =>
    op.move !== true &&
    losesMark(paragraph) &&
    !joinsNext.has(paragraph) &&
    (paragraph.pPrMark?.kind === "ins" ||
      paragraph.pPrMark?.kind === "moveTo") &&
    paragraph.pPrMark.deletion === undefined &&
    (!isOwnRecord(paragraph.pPrMark.info, revision) ||
      paragraph.sectionProperties !== undefined ||
      (beforeInsertedTable && paragraph === last.paragraph)) &&
    !(finalTakeover && paragraph === last.paragraph);
  if (
    affected.some(
      (paragraph) =>
        paragraph.pPrMark !== undefined &&
        !joinsNext.has(paragraph) &&
        !keepsDeletedMark(paragraph) &&
        !deletesAddedMark(paragraph),
    )
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      "Deletion would overwrite a paragraph review record.",
    );
  }
  // Revisions in the content resolve as a tracked range deletion over them
  // does; a content control is deleted whole, as one, keeping its identity,
  // and accepting the deletion removes it.
  const selectedSet = new Set(selected);
  const tracked: Paragraph[] = [];
  for (const paragraph of affected) {
    let next = paragraph;
    if (selectedSet.has(paragraph)) {
      const form = deletedContent(paragraph, op.revision, rangedComments);
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
    } else if (keepsDeletedMark(paragraph)) {
      // Already deleted: the earlier record stands.
    } else if (deletesAddedMark(paragraph) && next.pPrMark !== undefined) {
      next = {
        ...next,
        pPrMark: {
          ...next.pPrMark,
          deletion: { info: stampInfo(op.revision) },
        },
      };
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
      if (foldsIntoOtherSuggestion(paragraph, op.revision))
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          FOLD_INTO_OTHER_SUGGESTION_MESSAGE,
        );
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
  // What a pending deletion already holds whole, empty and with its mark
  // taken, a further one records nothing on: no reader or file would show it.
  // Where the direct deletion changes nothing either, as for an empty
  // paragraph alone in its container, recording nothing is that deletion.
  if (
    tracked.every((paragraph, index) => paragraph === affected[index]) &&
    !structurallyEqual(directBlocks, replacement([]).expected)
  )
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A tracked paragraph deletion would record nothing.",
    );
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
      const absorbed: ParagraphAbsorption[] = [];
      const modified = new Set<string>();
      let carried: ParagraphContent[] | undefined;
      let joining: string[] = [];
      for (const [index, paragraph] of marked.entries()) {
        const source = original[index];
        if (source !== undefined && joinsNext.has(source)) {
          carried = [...(carried ?? []), ...paragraph.content];
          removed.push(paragraph.paraId ?? "");
          joining.push(paragraph.paraId ?? "");
          continue;
        }
        const joined =
          carried === undefined
            ? paragraph
            : { ...paragraph, content: [...carried, ...paragraph.content] };
        if (carried !== undefined || changed[index] !== source)
          modified.add(joined.paraId ?? "");
        for (const id of joining)
          absorbed.push({ absorbed: id, survivor: joined.paraId ?? "" });
        joining = [];
        carried = undefined;
        blocks.push(joined);
      }
      const result = applyOps(document, [replacement(blocks)]);
      if (result.isErr())
        return refuse(result.error.reason, result.error.message);
      return Result.ok({
        ...result.value,
        touched: {
          modified: [...modified],
          inserted: [],
          removed,
          ...(absorbed.length === 0 ? {} : { absorbed }),
        },
      });
    }
    default: {
      const unreachable: never = fresh;
      return unreachable;
    }
  }
};
