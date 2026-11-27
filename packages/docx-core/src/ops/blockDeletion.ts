/** Whole-paragraph deletion within one block list, including its final paragraph. */
import { panic, Result } from "better-result";

import type {
  BlockContent,
  Document,
  DocumentBody,
  Paragraph,
  ParagraphContent,
  ParagraphPropertyChange,
  Table,
} from "../model/document";
import {
  blockListAt,
  blockListEndsInParagraph,
  type ParagraphLocation,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { newlyExhaustedCommentIds } from "./comments";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { freshenIdentities } from "./identity";
import { identityKeysIn, idKey, packageIdentityKeys } from "./ids";
import {
  asParagraphContent,
  childNodes,
  type InlineNode,
  isCommentAnchor,
  isCommentAnchorCarrierType,
  isEmptyContainer,
  rebuildNode,
} from "./leaves";
import { asMoveHalf } from "./moves";
import {
  isRangeBoundary,
  type ParagraphContentVerdict,
  paragraphHoldsContent,
  paragraphHoldsContentOnceAccepted,
  paragraphLength,
} from "./offsets";
import { pendingJoinsRunWordsInto } from "./pendingJoins";
import {
  blockFollower,
  isBlockBoundary,
  isBoundary,
  lastBlockPastBoundaries,
  withBoundariesOutsideDeletion,
} from "./rangeBoundaries";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  pendingParagraphMarkReason,
} from "./refusal";
import type { ApplyOps } from "./resolve";
import {
  foldedParagraphPropertyChange,
  isRemovedRevisionNode,
  paragraphPropertiesOf,
  paragraphPropertyChange,
  paragraphPropertyFoldConflict,
  reviewFieldsOf,
  sameParagraphProperties,
  stampInfo,
  takesOverOtherSuggestion,
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
  OP_STORIES,
  type OpStory,
  type ParagraphAbsorption,
  REVISION_DECISIONS,
  type RevisionStamp,
  SECTION_BOUNDARY_POLICIES,
  type SetSectionEndpointOp,
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
  | {
      kind: "deleted";
      content: ParagraphContent[];
      standing: InlineNode[];
      /** Whether the deletion marks a run, not only empty containers. */
      recordsContent: boolean;
    }
  | { kind: "unchanged"; standing: InlineNode[] }
  | { kind: "refused"; reason: DocumentOpRefusal["reason"] } => {
  // A comment's reference stands with its range's end, which the deletion
  // leaves standing wherever that end is; a point comment's reference is its
  // only anchor, which the deletion takes, as `w:del` around the run that
  // carries `w:commentReference`. Accepting it retires the comment, as
  // deleting the paragraph directly does; rejecting keeps it.
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
    takenReferences: pointReferenceIds(paragraph.content, rangedComments),
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
  const stands = standingIn(nodes, new Set(standing), created);
  // What the deletion records is judged once the boundaries left it: a
  // paragraph holding nothing else gets no record.
  const records = nodesWhere(nodes, created);
  if (records.length === 0) return { kind: "unchanged", standing: stands };
  return {
    kind: "deleted",
    content: asParagraphContent(nodes),
    standing: stands,
    recordsContent: records.some((record) =>
      childNodes(record)?.some((child) => !isEmptyContainer(child)),
    ),
  };
};

/** The nodes of `nodes` that `matches`, in document order, none inside another. */
const nodesWhere = (
  nodes: readonly InlineNode[],
  matches: (node: InlineNode) => boolean,
): InlineNode[] =>
  nodes.flatMap((node) =>
    matches(node) ? [node] : nodesWhere(childNodes(node) ?? [], matches),
  );

/** Whether `paragraph`'s own mark is pending deletion, as a move's source too. */
const markPendingDeletion = ({ pPrMark }: Paragraph): boolean =>
  pPrMark?.kind === "del" || pPrMark?.kind === "moveFrom";

/** Whether accepting every change takes `paragraph`'s own mark. */
export const markGoes = (paragraph: Paragraph): boolean =>
  markPendingDeletion(paragraph) || paragraph.pPrMark?.deletion !== undefined;

/** How many anchors of each comment stand anywhere in `value`. */
const commentAnchorCounts = (
  value: unknown,
  counts = new Map<number, number>(),
): Map<number, number> => {
  if (Array.isArray(value)) {
    for (const item of value) commentAnchorCounts(item, counts);
    return counts;
  }
  if (typeof value !== "object" || value === null) return counts;
  const type: unknown = Reflect.get(value, "type");
  const id: unknown = Reflect.get(value, "id");
  if (
    typeof type === "string" &&
    isCommentAnchorCarrierType(type) &&
    typeof id === "number"
  )
    counts.set(id, (counts.get(id) ?? 0) + 1);
  for (const item of Object.values(value)) commentAnchorCounts(item, counts);
  return counts;
};

type HeldCommentsOptions = {
  body: DocumentBody;
  paragraphs: readonly Paragraph[];
};

/**
 * The comments every anchor of which stands in `paragraphs`: removing them
 * leaves nothing of those comments, which retire with them, as a direct
 * deletion retires the comments it covers whole.
 */
const commentsHeldBy = ({
  body,
  paragraphs,
}: HeldCommentsOptions): ReadonlySet<number> => {
  const everywhere = commentAnchorCounts(body.content);
  return new Set(
    [...commentAnchorCounts(paragraphs)].flatMap(([id, count]) =>
      everywhere.get(id) === count ? [id] : [],
    ),
  );
};

/** Whether `node` is an anchor of one of the comments `held`. */
const isHeldAnchor = (node: InlineNode, held: ReadonlySet<number>): boolean =>
  isCommentAnchor(node) && held.has(node.id);

type AddedWholeOptions = {
  paragraph: Paragraph;
  owner: RevisionOwner;
  /** Comments wholly in the paragraphs removed: their anchors go too. */
  held: ReadonlySet<number>;
};

/**
 * Whether `owner`'s pending insertion added `paragraph` whole: its mark
 * and every word, with no one else's suggestion in it and no range boundary
 * standing in it, which removing the paragraph would lose, but the anchors
 * of comments `held` in what is removed.
 */
const addedWholeBy = ({ paragraph, owner, held }: AddedWholeOptions): boolean =>
  paragraph.pPrMark?.kind === "ins" &&
  paragraph.pPrMark.deletion === undefined &&
  isOwnRecord(paragraph.pPrMark.info, owner) &&
  paragraph.content.every(
    (item) =>
      (item.type === "insertion" && isOwnRecord(item.info, owner)) ||
      isHeldAnchor(item, held),
  ) &&
  !holdsRangeBoundary(paragraph.content, held) &&
  recordedOnlyBy(paragraph, owner);

type TakenOverBreakOptions = {
  /** The paragraph right before `paragraphs`. */
  broken: Paragraph;
  /** The paragraphs deleted, the last ones of their container. */
  paragraphs: readonly Paragraph[];
  owner: RevisionOwner;
  /** Comments wholly in `paragraphs`: rejecting the insertion retires them. */
  held: ReadonlySet<number>;
};

/**
 * The records whose rejection takes back `owner`'s insertion after the
 * container's last paragraph, if `paragraphs` are what it added: it gave the
 * paragraph before them (`broken`) a new break of its own, rejected into the
 * paragraph after it (`survivor: "first"`), each paragraph but the last an
 * added one, and the last one the words, with the former mark taken over.
 * Deleting them is rejecting those records, as deleting one's own insertion
 * removes it: `broken` keeps its identity and gets that mark and its
 * properties back. A record the former mark already carried stays on it.
 */
const takenOverBreakRecords = ({
  broken,
  paragraphs,
  owner,
  held,
}: TakenOverBreakOptions): number[] | undefined => {
  const mark = broken.pPrMark;
  const last = paragraphs.at(-1);
  if (
    last === undefined ||
    mark?.kind !== "ins" ||
    mark.deletion !== undefined ||
    mark.survivor !== "first" ||
    !isOwnRecord(mark.info, owner) ||
    broken.sectionProperties !== undefined ||
    !paragraphs
      .slice(0, -1)
      .every((paragraph) => addedWholeBy({ paragraph, owner, held }))
  )
    return undefined;
  const { pPrMark: formerMark, ...words } = last;
  if (
    (formerMark !== undefined &&
      (formerMark.kind !== "ins" || formerMark.deletion !== undefined)) ||
    !last.content.every(
      (item) =>
        (item.type === "insertion" && isOwnRecord(item.info, owner)) ||
        isHeldAnchor(item, held),
    ) ||
    holdsRangeBoundary(last.content, held) ||
    !recordedOnlyBy(words, owner)
  )
    return undefined;
  const ids = [mark.info.id];
  for (const paragraph of paragraphs) {
    if (paragraph !== last && paragraph.pPrMark !== undefined)
      ids.push(paragraph.pPrMark.info.id);
    for (const item of paragraph.content)
      if (item.type === "insertion") ids.push(item.info.id);
    for (const change of paragraph.propertyChanges ?? [])
      ids.push(change.info.id);
    if (paragraph.markPropertyChange !== undefined)
      ids.push(paragraph.markPropertyChange.info.id);
  }
  return ids;
};

/**
 * Whether a range boundary stands anywhere among `nodes`, but an anchor of a
 * comment `held` in what is removed.
 */
const holdsRangeBoundary = (
  nodes: readonly InlineNode[],
  held: ReadonlySet<number> = new Set(),
): boolean =>
  nodes.some(
    (node) =>
      (isBoundary(node) && !isHeldAnchor(node, held)) ||
      holdsRangeBoundary(childNodes(node) ?? [], held),
  );

/**
 * The ids of the comment references in `nodes` whose comment has no range
 * (`rangedComments` names those that do): a point comment, whose reference
 * is its only anchor.
 */
const pointReferenceIds = (
  nodes: readonly InlineNode[],
  rangedComments: ReadonlySet<number>,
): ReadonlySet<number> => {
  const ids = new Set<number>();
  const visit = (items: readonly InlineNode[]): void => {
    for (const node of items) {
      if (node.type === "commentReference" && !rangedComments.has(node.id))
        ids.add(node.id);
      visit(childNodes(node) ?? []);
    }
  };
  visit(nodes);
  return ids;
};

/** The ids of the comments whose range ends somewhere in `value`. */
export const rangeEndIdsIn = (
  value: unknown,
  ids: Set<number> = new Set(),
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
 * revision (`created`) or an earlier one holds goes with it. An anchor a
 * hyperlink, content control or transparent wrapper holds stands in what is
 * left of that container, as accepting the deletion leaves it there.
 */
const standingIn = (
  nodes: readonly InlineNode[],
  split: ReadonlySet<InlineNode>,
  created: (node: InlineNode) => boolean,
): InlineNode[] =>
  nodes.flatMap((node): InlineNode[] => {
    if (split.has(node) || isCommentAnchor(node)) return [node];
    if (created(node) || isRemovedRevisionNode(node)) return [];
    const inner = standingIn(childNodes(node) ?? [], split, created);
    return inner.length > 0 &&
      (node.type === "hyperlink" ||
        node.type === "inlineSdt" ||
        node.type === "inlineWrapper") &&
      holdsCommentAnchorIn(inner)
      ? [rebuildNode(node, inner)]
      : inner;
  });

/** Whether a comment anchor stands among `nodes`, at any depth. */
const holdsCommentAnchorIn = (nodes: readonly InlineNode[]): boolean =>
  nodes.some(
    (node) =>
      isCommentAnchor(node) || holdsCommentAnchorIn(childNodes(node) ?? []),
  );

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
  | { type: "blocked"; reason: "review"; review: "insertion" | "other" }
  | { type: "blocked"; reason: "section" | "structure" };

/** The first-cell path, including why it cannot receive paragraph content. */
export const paragraphTableRunInBoundary = (
  paragraph: Paragraph,
  table: Table,
): TableRunInBoundary => {
  if (paragraph.sectionProperties !== undefined)
    return { type: "blocked", reason: "section" };
  const tables: Table[] = [];
  let current = table;
  let review: "clear" | "insertion" | "other" = "clear";
  for (;;) {
    const row = current.rows.at(0),
      cell = row?.cells.at(0);
    if (row === undefined || cell === undefined)
      return { type: "blocked", reason: "structure" };
    if (
      row.structuralChange?.type === "tableRowDeletion" ||
      (cell.structuralChange !== undefined &&
        cell.structuralChange.type !== "tableCellInsertion")
    ) {
      review = "other";
    } else if (
      review === "clear" &&
      (row.structuralChange?.type === "tableRowInsertion" ||
        cell.structuralChange?.type === "tableCellInsertion")
    ) {
      review = "insertion";
    }
    tables.push(current);
    const target = cell.content.at(0);
    if (target?.type === "paragraph") {
      if (target.sectionProperties !== undefined)
        return { type: "blocked", reason: "section" };
      if (review !== "clear")
        return { type: "blocked", reason: "review", review };
      return { type: "receiver", target: { receiver: target, tables } };
    }
    if (target?.type !== "table")
      return { type: "blocked", reason: "structure" };
    current = target;
  }
};

/**
 * Whether the mark of `paragraph`, which `table` follows in `blocks`, can be
 * resolved away: the block list goes on ending with a paragraph, through a
 * block content control that ends it, and the table has a paragraph to
 * receive what runs in. A tracked deletion before a table is admitted only
 * where this holds, so each decision on it resolves.
 */
export const markBeforeTableResolves = ({
  blocks,
  boundary,
}: {
  blocks: readonly BlockContent[];
  boundary: TableRunInBoundary;
}): boolean => blockListEndsInParagraph(blocks) && boundary.type === "receiver";

type SectionEndpointRemovalOptions = {
  story: OpStory;
  location: ParagraphLocation;
};

type SectionEndpointRemoval = {
  ops: SetSectionEndpointOp[];
  /** The paragraph once its break is gone. */
  paragraph: Paragraph;
};

/**
 * The section break a body paragraph's mark carries, removed on its own:
 * the sections on either side become one under the later section's
 * properties. A paragraph with no break, or one that is not directly in the
 * main body, where no section ends, gets no operation.
 */
export const sectionEndpointRemoval = ({
  story,
  location: { list, paragraph },
}: SectionEndpointRemovalOptions): SectionEndpointRemoval => {
  if (
    story !== OP_STORIES.MAIN ||
    list.length > 0 ||
    paragraph.paraId === undefined ||
    paragraph.sectionProperties === undefined
  )
    return { ops: [], paragraph };
  const { sectionProperties, ...withoutBreak } = paragraph;
  return {
    ops: [
      {
        type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
        endpoint: { type: "paragraph", blockId: paragraph.paraId },
        expected: { type: "present", value: sectionProperties },
        properties: { type: "omitted" },
      },
    ],
    paragraph: withoutBreak,
  };
};

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
}):
  | { kind: "join" }
  | { kind: "remove" }
  | { kind: "run-in" }
  | { kind: "keep" }
  | { kind: "refuse"; reason: "untrackable" } => {
  switch (follows) {
    case "paragraph":
      return { kind: "join" };
    case "table":
      // An accepted deletion that leaves the paragraph empty is no run-in
      // source either: it goes without crossing the table, wherever the
      // story ends, as it does before a table that can receive words. A
      // block before it, or a paragraph still ending the container after the
      // table, anchors what its removal leaves.
      if (
        !tableRemovalIsTrackable &&
        (precededByBlock || containerStillEndsWithParagraph) &&
        !markWasAdded &&
        !holdsContent &&
        !needsFollowingCommentHost
      )
        return { kind: "remove" };
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
        };
      if (!holdsContent) {
        return needsFollowingCommentHost && followingTableCanReceiveRunIn
          ? { kind: "run-in" }
          : { kind: "remove" };
      }
      return followingTableCanReceiveRunIn
        ? { kind: "run-in" }
        : { kind: "keep" };
    case "block":
      // A resolved-away empty mark may go only while the container still
      // ends with a paragraph. Rejecting an added mark keeps standing words;
      // an accepted deletion may run its still-pending words into the block.
      if (!holdsContent)
        return containerStillEndsWithParagraph
          ? { kind: "remove" }
          : { kind: "keep" };
      if (markWasAdded) return { kind: "keep" };
      return !holdsContentOnceAccepted && followingBlockCanReceiveRunIn
        ? { kind: "run-in" }
        : { kind: "keep" };
    case "nothing":
      return !holdsContent && markWasAdded && containerStillEndsWithParagraph
        ? { kind: "remove" }
        : { kind: "keep" };
  }
};

type ExhaustedCommentsOptions = {
  before: Document;
  edit: DocumentEdit;
  applyOps: ApplyOps;
};

type ExhaustedCommentsRetirement = {
  edit: DocumentEdit;
  /** The paragraphs retiring the comments modified. */
  modified: readonly string[];
};

/**
 * `edit` with the comments it left nothing of retired, thread and all: those
 * that still covered something `before` it.
 */
const withExhaustedCommentsRetired = ({
  before,
  edit,
  applyOps,
}: ExhaustedCommentsOptions): Result<
  ExhaustedCommentsRetirement,
  DocumentOpRefusal
> => {
  const exhaustedCommentIds = newlyExhaustedCommentIds(before, edit.document);
  if (exhaustedCommentIds.length === 0)
    return Result.ok({ edit, modified: [] });
  const retired = applyOps(
    edit.document,
    exhaustedCommentIds.map((id) => ({
      type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
      id,
      scope: "thread" as const,
    })),
  );
  if (retired.isErr()) return Result.err(retired.error);
  return Result.ok({
    edit: {
      document: retired.value.document,
      inverse: [...retired.value.inverse, ...edit.inverse],
      touched: retired.value.touched,
    },
    modified: retired.value.touched.modified,
  });
};

type DirectRemovalReceiverOptions = {
  before: BlockContent | undefined;
  /** The blocks after the removed paragraphs, the one they stand before first. */
  after: readonly BlockContent[];
  kept: readonly ParagraphContent[];
};

/**
 * The paragraph that takes the boundaries of paragraphs a direct deletion
 * removes before a block, with its content once it holds them: the one
 * before them, at its end, else the block's first paragraph, at its start.
 * Past a block offering none, it is the first paragraph after it in the list,
 * as accepting their tracked deletion leaves them
 * ({@link planParagraphBoundaryRemoval}).
 */
const directRemovalReceiver = ({
  before,
  after,
  kept,
}: DirectRemovalReceiverOptions):
  { receiver: Paragraph; content: ParagraphContent[] } | undefined => {
  if (before?.type === "paragraph")
    return { receiver: before, content: [...before.content, ...kept] };
  const following = after.at(0);
  if (following === undefined) return undefined;
  const carrier: Paragraph = { type: "paragraph", content: [...kept] };
  let receiver: Paragraph | undefined;
  if (following.type === "table") {
    const boundary = paragraphTableRunInBoundary(carrier, following);
    if (boundary.type === "receiver") receiver = boundary.target.receiver;
  } else
    receiver =
      paragraphBlockRunInTarget(carrier, following)?.receiver ??
      after.find((block): block is Paragraph => block.type === "paragraph");
  return receiver === undefined
    ? undefined
    : { receiver, content: [...kept, ...receiver.content] };
};

/**
 * The index of the paragraph whose break a tracked deletion of the paragraphs
 * from `index` to their container's end retires: the one before them, past
 * the paragraphs pending deletions already take whole, words and mark, which
 * accepting them would leave it right before.
 */
const breakRetiredBefore = (
  list: readonly BlockContent[],
  index: number,
): number => {
  let candidate = index - 1;
  for (;;) {
    const block = list[candidate];
    if (
      block?.type !== "paragraph" ||
      !markPendingDeletion(block) ||
      block.sectionProperties !== undefined ||
      paragraphHoldsContentOnceAccepted(block)
    )
      break;
    candidate -= 1;
  }
  return candidate !== index - 1 && list[candidate]?.type === "paragraph"
    ? candidate
    : index - 1;
};

/**
 * A table pending deletion still blocks a tracked paragraph join, but it
 * does not supply properties to the closing paragraph. Look past such
 * tables for the paragraph that supplies them in the accepted document.
 */
const terminalPropertySource = (
  list: readonly BlockContent[],
  index: number,
): number => {
  let candidate = index;
  for (;;) {
    const block = list[candidate];
    if (
      block?.type !== "table" ||
      block.rows.length === 0 ||
      !block.rows.every(
        (row) => row.structuralChange?.type === "tableRowDeletion",
      )
    )
      return candidate;
    candidate = breakRetiredBefore(list, candidate);
  }
};

/**
 * Whether accepting leaves nothing of `blocks`, a container's last blocks:
 * paragraphs that keep no content once accepted, each before the last with
 * its break pending deletion, and no section break among them.
 */
const goesOnceAccepted = (blocks: readonly BlockContent[]): boolean =>
  blocks.length > 0 &&
  blocks.every(
    (block, index) =>
      block.type === "paragraph" &&
      block.sectionProperties === undefined &&
      !paragraphHoldsContentOnceAccepted(block) &&
      (index === blocks.length - 1 || markPendingDeletion(block)),
  );

const paragraphsIn = (blocks: readonly BlockContent[]): Paragraph[] =>
  blocks.filter((block): block is Paragraph => block.type === "paragraph");

/**
 * Whether the paragraph `list[index]` gives its properties to its container's
 * last one when a deletion retires its break: it holds words, its own or those
 * that pending mark deletions before it run into it. Not once its own mark is
 * pending deletion: the two already are one paragraph, with the last one's
 * properties.
 */
const givesProperties = (
  list: readonly BlockContent[],
  index: number,
): boolean => {
  const paragraph = list[index];
  return (
    paragraph?.type === "paragraph" &&
    !markPendingDeletion(paragraph) &&
    (paragraphLength(paragraph) > 0 || pendingJoinsRunWordsInto(list, index))
  );
};

const unchanged = (
  document: Document,
): Result<DocumentEdit, DocumentOpRefusal> =>
  Result.ok({
    document,
    inverse: [],
    touched: { modified: [], inserted: [], removed: [] },
  });

type FinalDeletionTakenFromOptions = {
  list: readonly BlockContent[];
  following: BlockContent | undefined;
  followerIndex: number;
  /** The last paragraph the deletion selects. */
  last: Paragraph;
};

/**
 * The container's last paragraph, right after `last`, when its pending
 * deletion took `last`'s mark, along with the property change that belongs
 * to that mark's deletion.
 */
const finalDeletionTakenFrom = ({
  list,
  following,
  followerIndex,
  last,
}: FinalDeletionTakenFromOptions) => {
  const mark = last.pPrMark;
  if (
    following?.type !== "paragraph" ||
    mark?.kind !== "del" ||
    blockFollower(list, followerIndex).block !== undefined ||
    paragraphHoldsContentOnceAccepted(following)
  )
    return undefined;
  const change = following.propertyChanges?.find(
    ({ resolvesWith }) => resolvesWith === mark.info.id,
  );
  return change === undefined ? undefined : { final: following, change };
};

type FinalDeletionTakenBackOptions = BlockDeletionOptions & {
  revision: RevisionStamp;
  /** The last paragraph whose deletion took the selection's last mark. */
  final: Paragraph;
  change: ParagraphPropertyChange;
};

/**
 * The tracked deletion of `op`'s paragraphs and `final` together, once the
 * property change `final`'s deletion gave it is rejected. The records the
 * earlier deletion made on the marks and words stay as they are. A
 * suggestion takes back no change of another suggestion or an ordinary one,
 * as it folds into none.
 */
const withFinalDeletionTakenBack = ({
  document,
  op,
  applyOps,
  revision,
  final,
  change,
}: FinalDeletionTakenBackOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  if (final.paraId === undefined)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      "The last paragraph has no id to delete it by.",
    );
  if (takesOverOtherSuggestion(change, revision))
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
      "A suggestion cannot take back another pending paragraph property change.",
    );
  const takenBack = applyOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: op.story,
      revisionIds: [change.info.id],
      decision: REVISION_DECISIONS.REJECT,
    },
  ]);
  if (takenBack.isErr())
    return refuse(takenBack.error.reason, takenBack.error.message);
  const deleted = deleteBlocks({
    document: takenBack.value.document,
    op: { ...op, blockIds: [...op.blockIds, final.paraId] },
    applyOps,
  });
  if (deleted.isErr()) return deleted;
  const { touched } = deleted.value;
  return Result.ok({
    ...deleted.value,
    inverse: [...deleted.value.inverse, ...takenBack.value.inverse],
    touched: {
      ...touched,
      modified: [
        ...new Set([...touched.modified, ...takenBack.value.touched.modified]),
      ],
    },
  });
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
  // Range boundaries the block list holds between the last paragraph and the
  // next one stand where the two meet once its mark goes, as resolving that
  // mark joins them: what stood before them stays before.
  const endsNoSection = found.every(
    ({ paragraph }) => paragraph.sectionProperties === undefined,
  );
  const follower = endsNoSection
    ? blockFollower(list, last.index)
    : { index: last.index + 1, block: list[last.index + 1], standing: [] };
  const { index: followerIndex, block: following } = follower;
  const standing = following?.type === "paragraph" ? follower.standing : [];
  const terminal = following === undefined;
  const endsWithParagraph = lastBlockPastBoundaries(list)?.type === "paragraph";
  // Paragraphs the same author's pending insertion added whole, with no one
  // else's suggestion in them, go outright before a paragraph or a table, as
  // deleting over one's own insertion removes it; no deletion of an addition
  // is recorded. Before a paragraph, an added mark that ends a section is not
  // taken back here: where its break goes is what rejecting the insertion
  // decides.
  const owner = op.revision;
  const beforeParagraph =
    following?.type === "paragraph" && op.move !== true && endsNoSection;
  const beforeTable = following?.type === "table" && endsWithParagraph;
  // A move carries its words, and their comments, to where it inserts them:
  // no comment retires with its source.
  const held =
    op.move === true
      ? new Set<number>()
      : commentsHeldBy({
          body,
          paragraphs: found.map(({ paragraph }) => paragraph),
        });
  if (
    owner !== undefined &&
    (beforeParagraph || beforeTable) &&
    found.every(({ paragraph }) => addedWholeBy({ paragraph, owner, held }))
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
    const retired = withExhaustedCommentsRetired({
      before: document,
      edit: removed.value,
      applyOps,
    });
    if (retired.isErr())
      return refuse(retired.error.reason, retired.error.message);
    return Result.ok({
      document: retired.value.edit.document,
      inverse: retired.value.edit.inverse,
      touched: {
        modified: retired.value.modified,
        inserted: [],
        removed: found.map(({ paragraph }) => paragraph.paraId ?? ""),
      },
    });
  }
  const blockBefore = list[first.index - 1];
  const broken = blockBefore?.type === "paragraph" ? blockBefore : undefined;
  const takenOver =
    owner !== undefined && terminal && op.move !== true && broken !== undefined
      ? takenOverBreakRecords({
          broken,
          paragraphs: found.map(({ paragraph }) => paragraph),
          owner,
          held,
        })
      : undefined;
  if (takenOver !== undefined && broken !== undefined) {
    const rejected = applyOps(document, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: op.story,
        revisionIds: takenOver,
        decision: REVISION_DECISIONS.REJECT,
      },
    ]);
    if (rejected.isErr())
      return refuse(rejected.error.reason, rejected.error.message);
    const { document: retracted, inverse, touched } = rejected.value;
    const survivor =
      storyParagraphs(storyBody(retracted, op.story)).find(
        ({ paragraph }) =>
          idKey(paragraph.paraId ?? "") === idKey(broken.paraId ?? ""),
      )?.paragraph ??
      panic("A taken-back insertion lost the paragraph it broke.");
    // Its own pending property changes are no part of the insertion: they
    // stay on `broken`, over the properties its mark gets back, as deleting
    // the paragraphs directly leaves its properties. Without any, the mark
    // it gets back reads as it stood before every change taken back, its
    // own pending change from the properties it had included (one the
    // paragraph `broken` was inserted after took over the story's last
    // mark): `broken` keeps its properties over that, as that pending
    // change. A taken-back paragraph with no such change records none.
    const { pPrMark: _addedMark, ...restyle } = reviewFieldsOf(broken);
    if (!broken.propertyChanges?.length) {
      const recorded = last.paragraph.propertyChanges?.[0];
      if (
        !recorded ||
        sameParagraphProperties(survivor.formatting, broken.formatting)
      )
        return Result.ok({ document: retracted, inverse, touched });
      restyle.propertyChanges = [
        paragraphPropertyChange(recorded.info, survivor, restyle.formatting),
      ];
    }
    const kept = applyOps(retracted, [
      {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
        story: op.story,
        blockId: survivor.paraId ?? "",
        expected: reviewFieldsOf(survivor),
        review:
          survivor.pPrMark === undefined
            ? restyle
            : { ...restyle, pPrMark: survivor.pPrMark },
      },
    ]);
    if (kept.isErr()) return refuse(kept.error.reason, kept.error.message);
    return Result.ok({
      document: kept.value.document,
      inverse: [...kept.value.inverse, ...inverse],
      touched,
    });
  }
  // The container's last paragraph whose pending deletion took this one's
  // mark, deleted its words and gave it this paragraph's properties: with a
  // paragraph before them whose break the two retire, deleting this one too
  // takes that property change back and deletes both together, as deleting
  // this paragraph first and the last one after does, and the last paragraph
  // takes that paragraph's properties. Back to the container's start or a
  // block, no break is retired: the run settles on its first paragraph, whose
  // properties the last one keeps, as it does deleted directly.
  const takenFinal = finalDeletionTakenFrom({
    list,
    following,
    followerIndex,
    last: last.paragraph,
  });
  const retiresBreakBefore =
    list[breakRetiredBefore(list, first.index)]?.type === "paragraph";
  if (
    op.revision !== undefined &&
    op.move !== true &&
    takenFinal !== undefined &&
    retiresBreakBefore
  )
    return withFinalDeletionTakenBack({
      document,
      op,
      applyOps,
      revision: op.revision,
      ...takenFinal,
    });
  // Tracked, a paragraph a table or another block follows loses its words
  // and its own mark, as the default editor records it; accepting removes
  // it, which needs what resolving that mark needs.
  const tableBoundary =
    following?.type === "table"
      ? paragraphTableRunInBoundary(last.paragraph, following)
      : undefined;
  // Before a table whose first row is itself pending review, or from a
  // paragraph whose mark ends a section, the words have nowhere to run in,
  // but the deletion takes them all: accepting leaves the paragraph empty,
  // and an empty paragraph goes without crossing the table, first in its
  // container or after another block, as resolving its mark already does.
  // A section break goes with its mark.
  const emptiedBeforeBlockedTable =
    tableBoundary?.type === "blocked" &&
    (tableBoundary.reason === "review" || tableBoundary.reason === "section") &&
    (endsWithParagraph || blockListEndsInParagraph(list));
  const trackedBeforeTable =
    op.revision !== undefined &&
    tableBoundary !== undefined &&
    (markBeforeTableResolves({ blocks: list, boundary: tableBoundary }) ||
      emptiedBeforeBlockedTable);
  const trackedBeforeBlock =
    op.revision !== undefined &&
    isOpaqueFollower(following) &&
    endsWithParagraph;
  // Directly, the breaks their marks carry go first, on their own, as
  // accepting their tracked deletion takes them; the paragraphs then go as
  // any others do.
  if (
    op.revision === undefined &&
    following !== undefined &&
    following.type !== "paragraph"
  ) {
    const breaks = found.filter(
      ({ paragraph }) => paragraph.sectionProperties !== undefined,
    );
    const ops = breaks.flatMap(
      (location) => sectionEndpointRemoval({ story: op.story, location }).ops,
    );
    if (breaks.length > 0 && ops.length === breaks.length) {
      const dropped = applyOps(document, ops);
      if (dropped.isErr())
        return refuse(dropped.error.reason, dropped.error.message);
      const removed = deleteBlocks({
        document: dropped.value.document,
        op,
        applyOps,
      });
      if (removed.isErr()) return removed;
      return Result.ok({
        ...removed.value,
        inverse: [...removed.value.inverse, ...dropped.value.inverse],
      });
    }
  }
  // Directly, paragraphs a table or another block follows, with no table
  // before them, go whole, as accepting their tracked deletion removes them:
  // the container still ends with a paragraph, and no two tables meet.
  const directBeforeBlock =
    op.revision === undefined &&
    following !== undefined &&
    following.type !== "paragraph" &&
    endsWithParagraph &&
    (first.index === 0 || list[first.index - 1]?.type === "paragraph") &&
    endsNoSection;
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
  const previousIndex =
    terminal && op.revision !== undefined
      ? breakRetiredBefore(list, first.index)
      : first.index - 1;
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
  /** `blocks` in place of the paragraphs `replaced`, through the next one. */
  const replacement = (
    blocks: readonly BlockContent[],
    replaced: readonly Paragraph[] = expected,
  ): DocumentOp =>
    following?.type === "paragraph" && standing.length > 0
      ? {
          // Boundaries between blocks are no paragraphs a replacement
          // addresses: the list is set whole, around the paragraph after.
          type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
          story: op.story,
          blockId: following.paraId ?? "",
          expected: list,
          blocks: [
            ...list.slice(0, list.indexOf(replaced.at(0) ?? following)),
            ...blocks,
            ...list.slice(followerIndex + 1),
          ],
        }
      : {
          type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
          story: op.story,
          expected: replaced,
          // With no boundaries standing, only paragraphs replace paragraphs.
          blocks: blocks.filter(
            (block): block is Paragraph => block.type === "paragraph",
          ),
          ...sectionBoundaries,
        };
  const finalReview = reviewFieldsOf(last.paragraph);
  const formattingSourceIndex =
    terminal && op.revision !== undefined
      ? terminalPropertySource(list, previousIndex)
      : previousIndex;
  const formattingSource = list[formattingSourceIndex];
  if (
    terminal &&
    formattingSource?.type === "paragraph" &&
    givesProperties(list, formattingSourceIndex)
  ) {
    const formatting = withMarkFormatting(
      paragraphPropertiesOf(formattingSource.formatting),
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
      kept.length === 0 && standing.length === 0
        ? following
        : {
            ...following,
            content: [...kept, ...standing, ...following.content],
          },
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
        after: list.slice(followerIndex),
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
  // The paragraph that ends its container keeps its place, so deleting one
  // with no words and none before it to join changes nothing.
  const changesNothing = structurallyEqual(directBlocks, expected);
  if (changesNothing && op.revision === undefined) return unchanged(document);
  directOps.push(replacement(directBlocks));
  const direct = applyOps(document, directOps);
  if (direct.isErr()) return refuse(direct.error.reason, direct.error.message);
  // A paragraph the direct deletion writes back is modified exactly when it
  // is not the block it replaces: the survivor, or the paragraph after the
  // deletion once it holds the boundaries the deleted paragraphs kept.
  const replaced = new Map(
    expected.map((block) => [idKey(block.paraId ?? ""), block]),
  );
  directModified.push(
    ...directBlocks
      .filter((block) => replaced.get(idKey(block.paraId ?? "")) !== block)
      .map(({ paraId }) => paraId ?? ""),
  );
  if (op.revision === undefined) {
    const retired = withExhaustedCommentsRetired({
      before: document,
      edit: direct.value,
      applyOps,
    });
    if (retired.isErr())
      return refuse(retired.error.reason, retired.error.message);
    // Anchoring the inverse must not report an unchanged boundary as modified.
    return Result.ok({
      ...retired.value.edit,
      touched: {
        modified: [...new Set([...retired.value.modified, ...directModified])],
        inserted: [],
        removed: affected
          .filter((paragraph) => !terminal || paragraph !== last.paragraph)
          .map(({ paraId }) => paraId ?? ""),
      },
    });
  }
  const revision = op.revision;
  // A range whose last break a pending deletion already takes, into
  // paragraphs that keep nothing once accepted up to its container's end,
  // ends that container once accepted, as a range at its end does: its
  // deletion retires the break before it too, as it would had those
  // paragraphs' deletions been accepted first. A range opening its container
  // has no break before it to retire.
  const retiredIndex =
    !terminal &&
    first.index > 0 &&
    markPendingDeletion(last.paragraph) &&
    goesOnceAccepted(
      list.slice(last.index + 1).filter((block) => !isBlockBoundary(block)),
    )
      ? breakRetiredBefore(list, first.index)
      : undefined;
  const retired = retiredIndex === undefined ? undefined : list[retiredIndex];
  // The paragraphs after the range, when its container's last one takes the
  // properties of the paragraph whose break is retired, as a deleted last
  // paragraph does.
  const trailing: Paragraph[] = [];
  let endFormatting = survivor.formatting;
  let propertySource = previous;
  let reviewed = affected;
  if (retiredIndex !== undefined && retired?.type === "paragraph") {
    const ending = lastBlockPastBoundaries(list);
    if (
      ending?.type === "paragraph" &&
      // The final paragraph's independent review keeps its own properties.
      // A deletion-owned change was already taken back above.
      (ending.propertyChanges?.length ?? 0) === 0 &&
      givesProperties(list, retiredIndex) &&
      !sameParagraphProperties(retired.formatting, ending.formatting)
    ) {
      trailing.push(...paragraphsIn(list.slice(last.index + 1)));
      propertySource = retired;
      endFormatting = withMarkFormatting(
        paragraphPropertiesOf(retired.formatting),
        ending.formatting,
      );
    }
    reviewed = [
      retired,
      ...paragraphsIn(list.slice(retiredIndex + 1, first.index)),
      ...selected,
      ...trailing,
    ];
  }
  const endsContainer = trailing.at(-1) ?? (terminal ? last.paragraph : null);
  // Before a block, accepting keeps the last paragraph when pending joins
  // run words into the deletion: its mark has nowhere to go, and it holds
  // those words. Directly, the paragraphs go whole and the words keep the
  // properties of the paragraph that holds them, so the last paragraph takes
  // those, as a pending change of its own.
  const joinedWordsHolder =
    trackedBeforeBlock && pendingJoinsRunWordsInto(list, first.index)
      ? list[first.index - 1]
      : undefined;
  const holderRestyle =
    joinedWordsHolder?.type === "paragraph"
      ? {
          formatting: withMarkFormatting(
            paragraphPropertiesOf(joinedWordsHolder.formatting),
            last.paragraph.formatting,
          ),
        }
      : undefined;
  // The paragraphs whose own mark the tracked deletion records deleted.
  const losesMark = (paragraph: Paragraph) =>
    paragraph !== endsContainer && !trailing.includes(paragraph);
  // A mark the same author's pending insertion or move added goes outright,
  // joining its paragraph to the next one, as deleting over one's own
  // insertion removes it; no deletion of an addition is recorded. Before a
  // table, the last paragraph has none to join and keeps the refusal. A mark
  // that ends a section is not taken back here: where its break goes is what
  // rejecting the insertion decides, so the deletion is recorded beside it
  // and that resolution moves the break.
  const joinsNext = new Set(
    reviewed.filter(
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
  // break goes either way, under the deletion that took it first. So does an
  // added mark whose later deletion already took it. A move's source records
  // its own paragraphs' marks moved, so one already taken is not; a
  // paragraph a pending deletion takes whole between the last paragraph and
  // the break the move retires keeps its record.
  const keepsDeletedMark = (paragraph: Paragraph) =>
    (markPendingDeletion(paragraph) ||
      paragraph.pPrMark?.deletion !== undefined) &&
    !(op.move === true && selected.includes(paragraph));
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
      (beforeInsertedTable && paragraph === last.paragraph));
  const overwrittenMark = reviewed.find(
    (paragraph) =>
      paragraph.pPrMark !== undefined &&
      !trailing.includes(paragraph) &&
      !joinsNext.has(paragraph) &&
      !keepsDeletedMark(paragraph) &&
      !deletesAddedMark(paragraph),
  )?.pPrMark;
  if (overwrittenMark !== undefined) {
    return refuse(
      pendingParagraphMarkReason(overwrittenMark),
      "Deletion would overwrite a paragraph review record.",
    );
  }
  // Revisions in the content resolve as a tracked range deletion over them
  // does; a content control is deleted whole, as one, keeping its identity,
  // and accepting the deletion removes it.
  const selectedSet = new Set(selected);
  const tracked: Paragraph[] = [];
  for (const paragraph of reviewed) {
    let next = paragraph;
    if (selectedSet.has(paragraph)) {
      const form = deletedContent(paragraph, op.revision, rangedComments);
      if (form.kind === "refused")
        return refuse(
          form.reason,
          "The selected paragraph content cannot be tracked.",
        );
      // A paragraph holding only empty containers, whose mark is already
      // pending deletion, has nothing left to record: a deletion of the empty
      // containers alone would be a change no reviewer sees.
      const nothingToRecord =
        form.kind === "deleted" &&
        !form.recordsContent &&
        keepsDeletedMark(paragraph);
      if (form.kind === "deleted" && !nothingToRecord)
        next = {
          ...paragraph,
          content:
            op.move === true
              ? asMoveHalf(form.content, op.revision.id, "from")
              : form.content,
        };
    }
    let restyle = paragraph === last.paragraph ? holderRestyle : undefined;
    if (joinsNext.has(paragraph)) {
      restyle = undefined;
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
    } else {
      if (losesMark(paragraph))
        next = {
          ...next,
          // A move's source marks its paragraph moved from, not deleted.
          pPrMark: {
            kind: op.move === true ? "moveFrom" : "del",
            info: stampInfo(op.revision),
          },
        };
      if (paragraph === endsContainer) restyle = { formatting: endFormatting };
    }
    if (
      restyle !== undefined &&
      !sameParagraphProperties(next.formatting, restyle.formatting)
    ) {
      const conflict = paragraphPropertyFoldConflict(paragraph, op.revision);
      if (conflict !== undefined)
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
          conflict,
        );
      const review = reviewFieldsOf(next);
      delete review.formatting;
      if (restyle.formatting !== undefined)
        review.formatting = restyle.formatting;
      review.propertyChanges = [
        foldedParagraphPropertyChange({
          paragraph,
          formatting: restyle.formatting,
          stamp: revision,
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
  // Where the direct deletion still changes something, as for a container's
  // last paragraph, which keeps its mark, holding only content a pending
  // deletion already takes, the suggestion records nothing a reviewer could
  // see or decide: a request that changes nothing.
  if (tracked.every((paragraph, index) => paragraph === reviewed[index])) {
    if (changesNothing) return unchanged(document);
    const pending = selected.find(keepsDeletedMark)?.pPrMark;
    if (pending !== undefined)
      return refuse(
        pendingParagraphMarkReason(pending),
        "The paragraph's pending deletion already owns the break.",
      );
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
      "A tracked paragraph deletion would record nothing.",
    );
  }
  const beforeKeys = new Set(identityKeysIn(reviewed));
  const fresh = freshenIdentities({
    before: reviewed,
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
      const next =
        following?.type === "paragraph" && trailing.length === 0
          ? [following]
          : [];
      // Several marks can start with the operation stamp. Link to the mark
      // that actually supplied the properties after those ids are distinct.
      const sourceMark = fresh.paragraphs.find(
        (_, index) => reviewed[index] === propertySource,
      )?.pPrMark;
      const ownerId =
        sourceMark?.kind === "del"
          ? sourceMark.info.id
          : sourceMark?.deletion?.info.id;
      const newOwner =
        ownerId !== undefined &&
        propertySource !== undefined &&
        !markGoes(propertySource)
          ? ownerId
          : undefined;
      const marked = [...fresh.paragraphs, ...next];
      const original = [...reviewed, ...next];
      const changed = [...tracked, ...next];
      // Each paragraph whose mark went joins the one after it, which keeps
      // its own identity and properties.
      const blocks: BlockContent[] = [];
      const removed: string[] = [];
      const absorbed: ParagraphAbsorption[] = [];
      const modified = new Set<string>();
      let carried: ParagraphContent[] | undefined;
      let joining: string[] = [];
      for (const [index, markedParagraph] of marked.entries()) {
        const source = original[index];
        const change = markedParagraph.propertyChanges?.at(0);
        const paragraph =
          newOwner !== undefined &&
          source === endsContainer &&
          !source.propertyChanges?.length &&
          change !== undefined
            ? {
                ...markedParagraph,
                propertyChanges: [{ ...change, resolvesWith: newOwner }],
              }
            : markedParagraph;
        // The boundaries standing before the following paragraph stay
        // between the blocks, or join it with a paragraph that joins it.
        if (source === following && standing.length > 0) {
          if (carried === undefined) blocks.push(...standing);
          else carried.push(...standing);
        }
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
      const result = applyOps(document, [replacement(blocks, original)]);
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
