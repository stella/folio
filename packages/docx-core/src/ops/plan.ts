import { panic, Result } from "better-result";
/**
 * Pure planning around tracked operations: how many new revision ids an
 * operation takes, and the operations a tracked deletion of a range is made
 * of once the author's own insertions and comment anchors are accounted for.
 */

import { type Document, MAX_REVISION_ID } from "../model/document";
import { type AppliedDocumentOp, applyDocumentOp, stampOf } from "./apply";
import {
  blockListAt,
  type ParagraphLocation,
  sameBlockList,
  storyParagraphs,
} from "./blocks";
import { structurallyEqual } from "./equality";
import {
  freshLeasedIds,
  ID_LEASE_KINDS,
  type IdLease,
  idLeaseRemaining,
} from "./idLease";
import {
  IDENTITY_SPACES,
  identityKeysIn,
  type IdentitySpace,
  idKey,
  packageIdentityKeyIndex,
  packageIdentityKeys,
  slotKey,
} from "./ids";
import {
  compareGaps,
  type Gap,
  isCommentAnchor,
  isEmptyContainer,
  isParagraphContent,
  type LeafSpan,
  leafSpans,
  zeroWidthLeavesAt,
} from "./leaves";
import { isRangeBoundary, paragraphLength } from "./offsets";
import {
  findStoryParagraph,
  locateParagraphInBlockList,
} from "./paragraphLocator";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  isAddedRevision,
  isRemovedRevisionNode,
  isTrackedWrapper,
  paragraphPropertiesOf,
} from "./review";
import { isOwnRecord, type RevisionOwner } from "./revisionOwner";
import { findStoryBody, sameStory, storyBody } from "./stories";
import {
  type CreateCommentOp,
  type DeleteRangeOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  EMPTY_PROPERTY_SETS,
  type InsertContentOp,
  type JoinBlocksOp,
  type NewIds,
  type OpStory,
  type ReplaceInlineOp,
  type RevisionStamp,
  SECTION_BOUNDARY_POLICIES,
  type SetParagraphPropsOp,
  type SetRunPropsOp,
  type SplitBlockOp,
  type TextPosition,
} from "./types";

/** The ids of one space among slot keys. */
const idsIn = (keys: readonly string[], space: IdentitySpace): number[] => {
  const prefix = `${space}:`;
  return keys.flatMap((key) =>
    key.startsWith(prefix) ? [Number(key.slice(prefix.length))] : [],
  );
};

/** The ids a package uses in one space. */
const usedIds = (document: Document, space: IdentitySpace): Set<number> =>
  new Set(idsIn(packageIdentityKeys(document.package), space));

/** The same operation with other new ids; one that takes none is returned as it is. */
const withNewIds = (op: DocumentOp, newIds: NewIds): DocumentOp => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_ROW:
    case DOCUMENT_OP_TYPES.DELETE_ROW:
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
    case DOCUMENT_OP_TYPES.DELETE_TABLE:
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
    case DOCUMENT_OP_TYPES.DELETE_COLUMN:
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
    case DOCUMENT_OP_TYPES.SPLIT_CELL:
    case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
      return { ...op, newIds };
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_TABLE:
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
    case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
    case DOCUMENT_OP_TYPES.SET_REVISION_INFO:
    case DOCUMENT_OP_TYPES.SET_CONTENT_CONTROL:
      return op;
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

/** Filter operation pools while retaining absent identity spaces. */
export const filterNewIds = (
  op: DocumentOp,
  keep: (space: IdentitySpace, id: number) => boolean,
): DocumentOp => {
  if (!("newIds" in op) || op.newIds === undefined) return op;
  const { revision, control } = op.newIds;
  const newIds: NewIds = {};
  if (revision !== undefined)
    newIds.revision = revision.filter((id) =>
      keep(IDENTITY_SPACES.REVISION, id),
    );
  if (control !== undefined)
    newIds.control = control.filter((id) => keep(IDENTITY_SPACES.CONTROL, id));
  return withNewIds(op, newIds);
};

type TrimAppliedNewIdsOptions = {
  before: Document;
  op: DocumentOp;
  applied: AppliedDocumentOp;
  story: OpStory;
};

/** Keep only supplied ids that the successful operation put in changed paragraphs. */
export const trimAppliedNewIds = ({
  before,
  op,
  applied,
  story,
}: TrimAppliedNewIdsOptions): DocumentOp => {
  if (!("newIds" in op) || op.newIds === undefined) return op;
  const touched = new Set(
    [...applied.touched.modified, ...applied.touched.inserted].map(idKey),
  );
  const body = storyBody(applied.document, story);
  const paragraphs = [...touched].flatMap((key) => {
    const location = findStoryParagraph(body, key);
    return location === undefined ? [] : [location.paragraph];
  });
  const identities = new Set(identityKeysIn(paragraphs));
  const revisions = new Set(applied.revisions);
  const known = packageIdentityKeyIndex(before.package);
  return filterNewIds(op, (space, id) => {
    const key = `${space}:${id}`;
    return (
      !known.has(key) &&
      (identities.has(key) ||
        (space === IDENTITY_SPACES.REVISION && revisions.has(id)))
    );
  });
};

/** Package-aware explicit pools for every semantic comment anchor cut, including replies. */
export const allocateCommentAnchorIds = (
  document: Document,
  op: CreateCommentOp,
  lease?: IdLease,
): Result<NewIds, DocumentOpRefusal> => {
  const keys = [
    ...packageIdentityKeys(document.package),
    ...identityKeysIn(op),
  ];
  const firstIn = (space: IdentitySpace): number => {
    let largest = 0;
    for (const id of idsIn(keys, space)) largest = Math.max(largest, id);
    return largest + 1;
  };
  const firstRevision = Math.max(
    firstIn(IDENTITY_SPACES.REVISION),
    op.comment.id + 1,
  );
  const firstControl = firstIn(IDENTITY_SPACES.CONTROL);
  const taken = new Set(keys);
  const first = {
    [IDENTITY_SPACES.REVISION]: firstRevision,
    [IDENTITY_SPACES.CONTROL]: firstControl,
  };
  const leaseKind = {
    [IDENTITY_SPACES.REVISION]: ID_LEASE_KINDS.REVISION,
    [IDENTITY_SPACES.CONTROL]: ID_LEASE_KINDS.CONTROL,
  };
  // A leased author's pools are the lowest leased ids no record carries.
  const pool = (space: IdentitySpace, count: number): number[] =>
    lease === undefined
      ? Array.from({ length: count }, (_, index) => first[space] + index)
      : freshLeasedIds({
          lease,
          kind: leaseKind[space],
          count,
          taken: (id) =>
            taken.has(`${space}:${id}`) ||
            (space === IDENTITY_SPACES.REVISION && id === op.comment.id),
        });
  const attempt = (counts: { revision: number; control: number }) =>
    applyDocumentOp(document, {
      ...op,
      newIds: {
        revision: pool(IDENTITY_SPACES.REVISION, counts.revision),
        control: pool(IDENTITY_SPACES.CONTROL, counts.control),
      },
    });
  const capacity =
    lease === undefined
      ? {
          revision: Math.max(0, MAX_REVISION_ID - firstRevision + 1),
          control: Math.max(0, MAX_REVISION_ID - firstControl + 1),
        }
      : {
          revision: idLeaseRemaining(lease, ID_LEASE_KINDS.REVISION),
          control: idLeaseRemaining(lease, ID_LEASE_KINDS.CONTROL),
        };
  let count = 0;
  for (;;) {
    const trial = attempt({
      revision: Math.min(count, capacity.revision),
      control: Math.min(count, capacity.control),
    });
    if (trial.isOk()) break;
    if (trial.error.reason !== DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS)
      return Result.err(trial.error);
    const next = count === 0 ? 1 : count * 2;
    if (
      next > MAX_DEMAND ||
      count >= Math.max(capacity.revision, capacity.control)
    )
      return Result.err(
        new DocumentOpRefusal({
          opType: op.type,
          reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
          message: "Comment anchor cuts exhaust the package identity space.",
        }),
      );
    count = next;
  }
  const counts = {
    revision: Math.min(count, capacity.revision),
    control: Math.min(count, capacity.control),
  };
  for (const space of [IDENTITY_SPACES.REVISION, IDENTITY_SPACES.CONTROL]) {
    let low = 0;
    let high = counts[space];
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const trial = attempt({ ...counts, [space]: middle });
      if (trial.isOk()) high = middle;
      else if (trial.error.reason === DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS)
        low = middle + 1;
      else return Result.err(trial.error);
    }
    counts[space] = low;
  }
  return Result.ok({
    revision: pool(IDENTITY_SPACES.REVISION, counts.revision),
    control: pool(IDENTITY_SPACES.CONTROL, counts.control),
  });
};

/** The largest id count {@link revisionIdDemand} searches before giving up. */
const MAX_DEMAND = 1 << 16;

/**
 * How many `newIds.revision` entries an operation takes on a document,
 * besides its stamp's id: the length of the shortest list of unused ids it
 * applies with. Content-control ids are supplied freely while counting. The
 * refusal is the one the operation meets whatever ids it is given.
 */
export const revisionIdDemand = (
  document: Document,
  op: DocumentOp,
): Result<number, DocumentOpRefusal> => {
  // DELETE_RANGE consumes candidate ids only when a tracked cut needs to
  // freshen a record. A successful empty-pool apply proves a full synthetic
  // control pool could not change the demand or make an unused id invalid.
  if (
    op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
    op.undefinedFields === undefined
  ) {
    const unallocated = withNewIds(op, { revision: [], control: [] });
    if (applyDocumentOp(document, unallocated).isOk()) return Result.ok(0);
  }
  const largest = (ids: ReadonlySet<number>, floor: number): number => {
    let out = floor;
    for (const id of ids) out = Math.max(out, id);
    return out;
  };
  // Above every id the package and the operation's own content carry, so none is taken.
  const carried = identityKeysIn(op);
  const first =
    largest(
      new Set([
        ...usedIds(document, IDENTITY_SPACES.REVISION),
        ...idsIn(carried, IDENTITY_SPACES.REVISION),
      ]),
      stampOf(op)?.id ?? 0,
    ) + 1;
  const pool = (count: number): number[] =>
    Array.from({ length: count }, (_, index) => first + index);
  const firstControl =
    largest(
      new Set([
        ...usedIds(document, IDENTITY_SPACES.CONTROL),
        ...idsIn(carried, IDENTITY_SPACES.CONTROL),
      ]),
      0,
    ) + 1;
  const control = Array.from(
    { length: 64 },
    (_, index) => firstControl + index,
  );
  const attempt = (count: number) =>
    applyDocumentOp(
      document,
      withNewIds(op, { revision: pool(count), control }),
    );
  const needsMore = (count: number): boolean | DocumentOpRefusal => {
    const applied = attempt(count);
    if (applied.isOk()) return false;
    return applied.error.reason === DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS
      ? true
      : applied.error;
  };
  const none = needsMore(0);
  if (none !== true) {
    return none === false ? Result.ok(0) : Result.err(none);
  }
  let low = 0;
  let high = 1;
  for (;;) {
    const outcome = needsMore(high);
    if (outcome === false) break;
    if (outcome !== true) return Result.err(outcome);
    if (high >= MAX_DEMAND || first + high > MAX_REVISION_ID) {
      return Result.err(
        new DocumentOpRefusal({
          message: "The operation takes more new ids than a package can hold.",
          reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
          opType: op.type,
        }),
      );
    }
    low = high;
    high *= 2;
  }
  // The shortest list lies in (low, high]: every longer list applies too.
  while (high - low > 1) {
    const middle = Math.floor((low + high) / 2);
    if (needsMore(middle) === false) {
      high = middle;
    } else {
      low = middle;
    }
  }
  return Result.ok(high);
};

/** How a leaf of a range to delete is deleted. */
type LeafPlan = "direct" | "tracked" | "untouched" | "anchor";

const leafPlan = (
  { node, ancestors }: LeafSpan,
  owner: RevisionOwner,
): LeafPlan => {
  if (ancestors.some(isRemovedRevisionNode)) return "untouched";
  if (isCommentAnchor(node)) return "anchor";
  const own = ancestors.some(
    (ancestor) =>
      isAddedRevision(ancestor) &&
      isTrackedWrapper(ancestor) &&
      isOwnRecord(ancestor.info, owner),
  );
  return own ? "direct" : "tracked";
};

type Segment = {
  plan: "direct" | "tracked";
  from: Gap;
  to: Gap;
  records: boolean;
};

const laterGap = (left: Gap, right: Gap): Gap =>
  compareGaps(left, right) >= 0 ? left : right;
const earlierGap = (left: Gap, right: Gap): Gap =>
  compareGaps(left, right) <= 0 ? left : right;

/**
 * The stretches of a range to delete directly and with tracking, in document
 * order. A retained range or comment anchor ends a stretch; a tracked stretch runs across
 * deleted content, which it leaves alone, and one holding nothing else is
 * dropped.
 */
type ReplacementDeletionSegmentsOptions = {
  spans: readonly LeafSpan[];
  from: Gap;
  to: Gap;
  mode:
    | {
        type: "editing";
        /**
         * Whether a stretch holding only empty containers is dropped, as a
         * tracked one is. A direct edit drops it when its tracked counterpart
         * records something else and so keeps the containers. (default: false)
         */
        keepsLoneEmptyContainers?: boolean;
      }
    | { type: "suggesting"; owner: RevisionOwner };
};

export const replacementDeletionSegments = ({
  spans,
  from,
  to,
  mode,
}: ReplacementDeletionSegmentsOptions): Segment[] => {
  const keepsLoneEmptyContainers =
    mode.type === "editing" && mode.keepsLoneEmptyContainers === true;
  const out: Segment[] = [];
  let open: Segment | undefined;
  for (const span of spans) {
    const start = laterGap(span.before, from);
    const end = earlierGap(span.after, to);
    if (compareGaps(start, end) >= 0) continue;
    if (isParagraphContent(span.node) && isRangeBoundary(span.node)) {
      open = undefined;
      continue;
    }
    let plan: LeafPlan = isCommentAnchor(span.node) ? "anchor" : "direct";
    if (mode.type === "suggesting") plan = leafPlan(span, mode.owner);
    else if (span.ancestors.some(isRemovedRevisionNode)) plan = "untouched";
    if (plan === "anchor") {
      open = undefined;
      continue;
    }
    const segmentPlan = plan === "untouched" ? "tracked" : plan;
    // An empty container marks nothing tracked: a stretch holding nothing
    // else is dropped, as one over deleted content is. A suggestion drops it
    // inside the author's own insertion too, so that a direct edit, which
    // knows no author, can keep it the same way wherever it stands.
    const records =
      plan !== "untouched" &&
      !(
        (mode.type === "suggesting" || keepsLoneEmptyContainers) &&
        isEmptyContainer(span.node)
      );
    if (
      open !== undefined &&
      open.plan === segmentPlan &&
      compareGaps(open.to, start) === 0
    ) {
      open.to = end;
      open.records ||= records;
      continue;
    }
    open = { plan: segmentPlan, from: start, to: end, records };
    out.push(open);
  }
  return out.filter(({ records }) => records);
};

/**
 * The leaves of a range that its direct deletion leaves standing, in
 * document order: each one no stretch of {@link replacementDeletionSegments}
 * removes, such as a range boundary, a comment anchor or content already on
 * its way out. Text replacing the range goes in after the last of them.
 */
export const leavesKeptByDirectDeletion = ({
  spans,
  from,
  to,
  keepsLoneEmptyContainers = false,
}: Omit<ReplacementDeletionSegmentsOptions, "mode"> & {
  keepsLoneEmptyContainers?: boolean;
}): LeafSpan[] => {
  const segments = replacementDeletionSegments({
    spans,
    from,
    to,
    mode: { type: "editing", keepsLoneEmptyContainers },
  });
  return spans.filter((span) => {
    const start = laterGap(span.before, from);
    const end = earlierGap(span.after, to);
    return (
      compareGaps(start, end) < 0 &&
      !segments.some(
        (segment) =>
          compareGaps(segment.from, start) <= 0 &&
          compareGaps(end, segment.to) <= 0,
      )
    );
  });
};

/** The range a tracked deletion covers, its stamp, and the ids its later pieces take. */
export type PlanTrackedDeletionOptions = {
  from: TextPosition;
  to: TextPosition;
  revision: RevisionStamp;
  /** Ids for the stamps of the tracked deletions after the first, and for the records they cut. */
  newIds?: NewIds;
};

const positionOf = (at: TextPosition, gap: Gap): TextPosition => ({
  story: at.story,
  blockId: at.blockId,
  offset: gap.offset,
  zeroWidthBefore: gap.zeroWidthBefore,
});

const outOfIds = (): Result<never, DocumentOpRefusal> =>
  Result.err(
    new DocumentOpRefusal({
      message: "The planned deletions need more new revision ids.",
      reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
      opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
    }),
  );

/** Operations whose stamps the range planners allocate from one shared pool. */
type PlannedReviewOp =
  | DeleteRangeOp
  | InsertContentOp
  | JoinBlocksOp
  | SplitBlockOp
  | SetParagraphPropsOp
  | SetRunPropsOp
  | ReplaceInlineOp;

type TrackedPlanOptions = {
  document: Document;
  revision: RevisionStamp;
  newIds: NewIds | undefined;
  /**
   * Count every piece's new ids before applying it, as planning did before
   * it applied a piece with none first. Only its reference in tests.
   */
  countEveryPiece?: boolean;
  /**
   * The record a deletion carrying the edit's own id is stamped as, when it
   * continues the author's earlier deletion: the plan applies each piece as
   * it will be written, so whether the deletion makes a record of its own,
   * and so which id later pieces take, holds for the ops it gives.
   */
  deletionRevision?: RevisionStamp;
};

/** The story a planned piece edits. */
const storyOf = (input: PlannedReviewOp): OpStory => {
  switch (input.type) {
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return input.from.story;
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return input.at.story;
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
      return input.story;
    default: {
      const unreachable: never = input;
      return unreachable;
    }
  }
};

/** Whether an operation, its unused ids trimmed, still names a new id. */
const takesNewIds = (op: DocumentOp): boolean =>
  "newIds" in op &&
  op.newIds !== undefined &&
  ((op.newIds.revision?.length ?? 0) > 0 ||
    (op.newIds.control?.length ?? 0) > 0);

type PlannedParagraphJoin = Omit<
  JoinBlocksOp,
  "revision" | "newIds" | "survivor"
> & { revision?: never; newIds?: never; survivor?: never };
type TrackedParagraphJoinKind = "tracked-mark" | "cancelled-added-mark";
type TrackedParagraphJoinResult = Result<
  TrackedParagraphJoinKind,
  DocumentOpRefusal
>;
type TrackedPlan = {
  append: (
    input: Exclude<PlannedReviewOp, JoinBlocksOp>,
  ) => Result<void, DocumentOpRefusal>;
  join: (input: PlannedParagraphJoin) => TrackedParagraphJoinResult;
  document: () => Document;
  ops: DocumentOp[];
};

type AvailableIdsOptions = {
  document: Document;
  space: IdentitySpace;
  candidates: readonly number[] | undefined;
};

const availableIds = ({
  document,
  space,
  candidates,
}: AvailableIdsOptions): number[] => {
  if (!candidates?.length) return [];
  const used = packageIdentityKeyIndex(document.package);
  return candidates.filter((id) => !used.has(slotKey({ space, id })));
};

/** Shared allocation and simulation for all pieces of one planned edit. */
export const createTrackedPlan = ({
  document,
  revision,
  newIds,
  countEveryPiece = false,
  deletionRevision,
}: TrackedPlanOptions): TrackedPlan => {
  const pool = availableIds({
    document,
    space: IDENTITY_SPACES.REVISION,
    candidates: [...new Set(newIds?.revision)],
  }).filter((id) => id !== revision.id);
  let taken = 0;
  let remaining = pool.length;
  let stampUsed = false;
  let current = document;
  const ops: DocumentOp[] = [];
  const take = (count: number) => {
    if (taken + count > remaining) return undefined;
    const ids = pool.slice(taken, taken + count);
    taken += count;
    return ids;
  };
  const append = (input: PlannedReviewOp): Result<void, DocumentOpRefusal> => {
    let op: DocumentOp = input;
    if (
      input.type !== DOCUMENT_OP_TYPES.REPLACE_INLINE &&
      input.revision !== undefined
    ) {
      // Source cuts consume the caller's pool in the same order as editing.
      // Additional mode-owned stamps use its other end and cannot displace them.
      let stampId = revision.id;
      let stamp = revision;
      if (stampUsed) {
        if (remaining <= taken) return outOfIds();
        remaining -= 1;
        stampId =
          pool.at(remaining) ??
          panic("A reserved stamp has an allocated pool entry.");
      } else if (
        input.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
        deletionRevision !== undefined
      )
        stamp = deletionRevision;
      op = { ...input, revision: { ...stamp, id: stampId } };
    }
    const control = availableIds({
      document: current,
      space: IDENTITY_SPACES.CONTROL,
      candidates: newIds?.control,
    });
    const story = storyOf(input);
    // Most edits take no new id: applied with none, the edit is planned as
    // it is. It took none when it used none of the content-control ids
    // either, so counting would have found it needs none.
    const unallocated = withNewIds(op, { revision: [], control });
    const attempt = countEveryPiece
      ? undefined
      : applyDocumentOp(current, unallocated);
    let applied: Result<AppliedDocumentOp, DocumentOpRefusal> | undefined;
    let trimmed: DocumentOp | undefined;
    if (attempt?.isOk() === true) {
      trimmed = trimAppliedNewIds({
        before: current,
        op: unallocated,
        applied: attempt.value,
        story,
      });
      if (!takesNewIds(trimmed)) applied = attempt;
    }
    if (applied === undefined) {
      trimmed = undefined;
      const demand = revisionIdDemand(current, op);
      if (demand.isErr()) return Result.err(demand.error);
      const ids = take(demand.value);
      if (ids === undefined) return outOfIds();
      op = withNewIds(op, { revision: ids, control });
      applied = applyDocumentOp(current, op);
    } else {
      op = unallocated;
    }
    if (applied.isErr()) return Result.err(applied.error);
    stampUsed ||=
      stampOf(input) !== undefined && applied.value.revisions.length > 0;
    const before = current;
    current = applied.value.document;
    ops.push(
      trimmed ??
        trimAppliedNewIds({ before, op, applied: applied.value, story }),
    );
    return Result.ok(undefined);
  };
  const getDocument = () => current;
  return {
    append,
    join: (join) =>
      appendTrackedParagraphJoin({ append, getDocument, join, revision }),
    document: getDocument,
    ops,
  };
};

type AppendTrackedParagraphJoinOptions = {
  append: (input: PlannedReviewOp) => Result<void, DocumentOpRefusal>;
  getDocument: () => Document;
  join: PlannedParagraphJoin;
  revision: RevisionStamp;
};

/** Cancel an own added mark physically, leaving property changes to its planner. */
const appendTrackedParagraphJoin = ({
  append,
  getDocument,
  join,
  revision,
}: AppendTrackedParagraphJoinOptions): TrackedParagraphJoinResult => {
  const body = storyBody(getDocument(), join.story);
  const leading = findStoryParagraph(body, idKey(join.blockId))?.paragraph;
  const mark = leading?.pPrMark;
  const cancelsAddedMark =
    (mark?.kind === "ins" || mark?.kind === "moveTo") &&
    isOwnRecord(mark.info, revision);
  const before = findStoryParagraph(body, idKey(join.nextBlockId))?.paragraph;
  const joined = append(cancelsAddedMark ? join : { ...join, revision });
  if (joined.isErr()) return Result.err(joined.error);
  if (!cancelsAddedMark) return Result.ok("tracked-mark");
  const after = findStoryParagraph(
    storyBody(getDocument(), join.story),
    idKey(join.nextBlockId),
  )?.paragraph;
  if (!before || !after) panic("An own-mark join lost its trailing paragraph.");
  // A physical join takes the leading properties. Restore the survivor's
  // original properties before the planner records its desired properties
  // as a suggestion; otherwise reject would retain the cancelled mark's style.
  const restored = append({
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: join.story,
    blockId: join.nextBlockId,
    patch: Object.fromEntries([
      ...Object.keys(after.formatting ?? {}).map((key) => [key, null]),
      ...Object.entries(before.formatting ?? {}),
    ]),
    whenEmpty:
      before.formatting === undefined
        ? EMPTY_PROPERTY_SETS.OMIT
        : EMPTY_PROPERTY_SETS.KEEP,
  });
  return restored.isErr()
    ? Result.err(restored.error)
    : Result.ok("cancelled-added-mark");
};

/** A text range crosses containers without removing their geometry or boundary marks. */
export const selectedParagraphRuns = (
  document: Document,
  from: TextPosition,
  to: TextPosition,
): Result<ParagraphLocation[][], DocumentOpRefusal> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(
      new DocumentOpRefusal({
        reason,
        message,
        opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
      }),
    );
  if (!sameStory(from.story, to.story))
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A text range stays in one story.",
    );
  const body = storyBody(document, from.story);
  const missingEndpoint = () =>
    refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "A range endpoint does not exist.",
    );
  const reversedRange = () =>
    refuse(
      DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
      "Range endpoints must be in document order.",
    );
  const start = idKey(from.blockId);
  const end = idKey(to.blockId);
  if (start !== "" && end !== "") {
    const left = findStoryParagraph(body, start);
    const right = start === end ? left : findStoryParagraph(body, end);
    if (left === undefined || right === undefined) return missingEndpoint();
    if (left === right) return Result.ok([[left]]);
    if (sameBlockList(left.list, right.list)) {
      if (left.index > right.index) return reversedRange();
      const list = blockListAt(body.content, left.list);
      const run: ParagraphLocation[] = [];
      for (let index = left.index; index <= right.index; index += 1) {
        const paragraph = list[index];
        if (paragraph?.type !== "paragraph") break;
        run.push({ list: left.list, index, paragraph });
      }
      if (run.length === right.index - left.index + 1) return Result.ok([run]);
    }
  }
  const locations = storyParagraphs(body);
  const first = locations.findIndex(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === start,
  );
  const last = locations.findIndex(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === end,
  );
  if (first < 0 || last < 0) return missingEndpoint();
  if (first > last) return reversedRange();
  const runs: ParagraphLocation[][] = [];
  let previous: ParagraphLocation | undefined;
  for (const location of locations.slice(first, last + 1)) {
    const current = runs.at(-1);
    if (
      current &&
      previous &&
      sameBlockList(previous.list, location.list) &&
      location.index === previous.index + 1
    )
      current.push(location);
    else runs.push([location]);
    previous = location;
  }
  return Result.ok(runs);
};

type PlanDeletionOptions = {
  document: Document;
  options: PlanTrackedDeletionOptions;
  plan: TrackedPlan;
};

/** Validate and segment each paragraph before planning any edits. */
export const appendTrackedDeletion = ({
  document,
  options,
  plan,
}: PlanDeletionOptions): Result<void, DocumentOpRefusal> => {
  const { from, to, revision } = options;
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(
      new DocumentOpRefusal({
        reason,
        message,
        opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
      }),
    );
  const partitioned = selectedParagraphRuns(document, from, to);
  if (partitioned.isErr()) return Result.err(partitioned.error);
  if (partitioned.value.length > 1) {
    for (const run of partitioned.value.toReversed()) {
      const first = run.at(0)?.paragraph;
      const last = run.at(-1)?.paragraph;
      if (!first || !last)
        panic("A selected paragraph run must contain its endpoints.");
      const start =
        idKey(first.paraId ?? "") === idKey(from.blockId)
          ? from
          : {
              story: from.story,
              blockId: first.paraId ?? "",
              offset: 0,
              zeroWidthBefore: 0,
            };
      const end =
        idKey(last.paraId ?? "") === idKey(to.blockId)
          ? to
          : {
              story: to.story,
              blockId: last.paraId ?? "",
              offset: paragraphLength(last),
              zeroWidthBefore: zeroWidthLeavesAt(
                last.content,
                paragraphLength(last),
              ).length,
            };
      const appended = appendTrackedDeletion({
        document: plan.document(),
        options: { ...options, from: start, to: end },
        plan,
      });
      if (appended.isErr()) return appended;
    }
    return Result.ok(undefined);
  }
  const body = findStoryBody(document, from.story);
  if (!body)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The story does not exist.",
    );
  const run = partitioned.value.at(0);
  const first = run?.at(0);
  const last = run?.at(-1);
  if (first === undefined || last === undefined)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "A range endpoint does not exist.",
    );
  if (first.index > last.index)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
      "Range endpoints must be in document order.",
    );
  const list = blockListAt(body.content, first.list);
  const ranges = [];
  for (let index = first.index; index <= last.index; index += 1) {
    const paragraph = list[index];
    if (paragraph?.type !== "paragraph")
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        "A planned range cannot cross a non-paragraph block.",
      );
    const at = {
      story: from.story,
      blockId: paragraph.paraId ?? "",
      offset: 0,
    };
    const start = index === first.index ? from : { ...at, zeroWidthBefore: 0 };
    const end =
      index === last.index
        ? to
        : {
            ...at,
            offset: paragraphLength(paragraph),
            zeroWidthBefore: zeroWidthLeavesAt(
              paragraph.content,
              paragraphLength(paragraph),
            ).length,
          };
    const check = revisionIdDemand(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: start,
      to: end,
    });
    if (check.isErr()) return Result.err(check.error);
    const content = paragraph.content;
    const fromGap = {
      offset: start.offset,
      zeroWidthBefore:
        start.zeroWidthBefore ??
        zeroWidthLeavesAt(content, start.offset).length,
    };
    const toGap = {
      offset: end.offset,
      zeroWidthBefore: end.zeroWidthBefore ?? 0,
    };
    const segments = replacementDeletionSegments({
      spans: leafSpans(content),
      from: fromGap,
      to: toGap,
      mode: { type: "suggesting", owner: revision },
    });
    const only = segments.at(0);
    const reusesPreflight =
      first.index === last.index &&
      segments.length === 1 &&
      only !== undefined &&
      start.story === end.story &&
      start.blockId === at.blockId &&
      end.blockId === at.blockId &&
      compareGaps(only.from, fromGap) === 0 &&
      compareGaps(only.to, toGap) === 0;
    ranges.push({ at, segments, reusesPreflight });
  }
  // Direct simulation gives the surviving paragraph's authored properties,
  // independent of tracked-deleted content that still occupies offsets.
  let direct = document;
  for (const { at, segments, reusesPreflight } of ranges.toReversed()) {
    for (const segment of segments.toReversed()) {
      const op = {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: positionOf(at, segment.from),
        to: positionOf(at, segment.to),
      } satisfies DeleteRangeOp;
      const appended = plan.append(
        segment.plan === "tracked" ? { ...op, revision } : op,
      );
      if (appended.isErr()) return appended;
      // The preflight already applied this exact direct deletion to the
      // original document. A direct delete does not consume new IDs, and
      // this single-paragraph plan has no join that needs its output.
      if (reusesPreflight) return Result.ok(undefined);
      const allocated = plan.ops.at(-1);
      const simulated = applyDocumentOp(direct, {
        ...op,
        ...(allocated?.type === DOCUMENT_OP_TYPES.DELETE_RANGE
          ? { newIds: allocated.newIds }
          : {}),
      });
      if (simulated.isErr()) return Result.err(simulated.error);
      direct = simulated.value.document;
    }
  }
  if (first.index === last.index) return Result.ok(undefined);
  for (let index = ranges.length - 2; index >= 0; index -= 1) {
    const leading = ranges[index];
    if (leading === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        "A planned join has no endpoint.",
      );
    const currentBody = storyBody(plan.document(), from.story);
    const currentLeading = locateParagraphInBlockList(
      currentBody,
      leading.at.blockId,
    );
    if (currentLeading === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        "A planned join lost its leading paragraph.",
      );
    const mark = currentLeading.paragraph.pPrMark;
    if (mark?.kind === "del" || mark?.kind === "moveFrom") continue;
    const next = currentLeading.blocks.at(currentLeading.index + 1);
    if (next?.type !== "paragraph")
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
        "A planned join has no following paragraph.",
      );
    const join = {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: from.story,
      blockId: leading.at.blockId,
      nextBlockId: next.paraId ?? "",
      sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE,
    } satisfies JoinBlocksOp;
    const appended = plan.join(join);
    if (appended.isErr()) return Result.err(appended.error);
    const directBody = storyBody(direct, from.story);
    const directLeading = locateParagraphInBlockList(
      directBody,
      leading.at.blockId,
    );
    if (directLeading === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        "A simulated join lost its leading paragraph.",
      );
    const directNext = directLeading.blocks.at(directLeading.index + 1);
    if (directNext?.type !== "paragraph")
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
        "A simulated join has no following paragraph.",
      );
    const simulated = applyDocumentOp(direct, {
      ...join,
      nextBlockId: directNext.paraId ?? "",
    });
    if (simulated.isErr()) return Result.err(simulated.error);
    direct = simulated.value.document;
  }
  const survivor = locateParagraphInBlockList(
    storyBody(direct, from.story),
    to.blockId,
  );
  const trackedSurvivor = locateParagraphInBlockList(
    storyBody(plan.document(), from.story),
    to.blockId,
  );
  if (survivor === undefined || trackedSurvivor === undefined)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The planned survivor does not exist.",
    );
  const current = paragraphPropertiesOf(trackedSurvivor.paragraph.formatting);
  const desired = paragraphPropertiesOf(survivor.paragraph.formatting);
  if (structurallyEqual(current, desired)) return Result.ok(undefined);
  const patch = Object.fromEntries([
    ...Object.keys(current ?? {}).map((key) => [key, null]),
    ...Object.entries(desired ?? {}),
  ]);
  return plan.append({
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: from.story,
    blockId: to.blockId,
    patch,
    whenEmpty:
      survivor.paragraph.formatting === undefined
        ? EMPTY_PROPERTY_SETS.OMIT
        : EMPTY_PROPERTY_SETS.KEEP,
    revision,
  });
};

/**
 * Plan a same-list range as inline deletions followed by tracked paragraph
 * joins. Own insertions are retracted directly; comment anchors remain.
 * Inline ranges run in reverse order, then joins run from last to first.
 * All physical revision ids come from one pool for the complete batch.
 */
export const planTrackedDeletion = (
  document: Document,
  options: PlanTrackedDeletionOptions,
): Result<DocumentOp[], DocumentOpRefusal> => {
  const plan = createTrackedPlan({
    document,
    revision: options.revision,
    newIds: options.newIds,
  });
  const appended = appendTrackedDeletion({ document, options, plan });
  return appended.isErr() ? Result.err(appended.error) : Result.ok(plan.ops);
};
