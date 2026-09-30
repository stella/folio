/**
 * Pure planning around tracked operations: how many new revision ids an
 * operation takes, and the operations a tracked deletion of a range is made
 * of once the author's own insertions and comment anchors are accounted for.
 */

import { Result } from "better-result";

import { type Document, MAX_REVISION_ID } from "../model/document";
import { applyDocumentOp, stampOf } from "./apply";
import { blockListAt, sameBlockList, storyBody, storyParagraphs } from "./blocks";
import { structurallyEqual } from "./equality";
import {
  IDENTITY_SPACES,
  type IdentitySpace,
  identityKeysIn,
  idKey,
  packageIdentityKeys,
} from "./ids";
import { paragraphLength } from "./offsets";
import { compareGaps, type Gap, type LeafSpan, leafSpans, zeroWidthLeavesAt } from "./leaves";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { isRemovedRevisionNode, isTrackedWrapper, paragraphPropertiesOf } from "./review";
import {
  DOCUMENT_OP_TYPES,
  type DeleteRangeOp,
  type InsertContentOp,
  type JoinBlocksOp,
  type SetParagraphPropsOp,
  type SplitBlockOp,
  EMPTY_PROPERTY_SETS,
  type DocumentOp,
  type NewIds,
  type RevisionStamp,
  type TextPosition,
} from "./types";

/** The ids of one space among slot keys. */
const idsIn = (keys: readonly string[], space: IdentitySpace): number[] => {
  const prefix = `${space}:`;
  return keys.flatMap((key) => (key.startsWith(prefix) ? [Number(key.slice(prefix.length))] : []));
};

/** The ids a package uses in one space. */
const usedIds = (document: Document, space: IdentitySpace): Set<number> =>
  new Set(idsIn(packageIdentityKeys(document.package), space));

/** The same operation with other new ids; one that takes none is returned as it is. */
const withNewIds = (op: DocumentOp, newIds: NewIds): DocumentOp => {
  switch (op.type) {
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
      return { ...op, newIds };
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
      return op;
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
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
  const control = Array.from({ length: 64 }, (_, index) => firstControl + index);
  const attempt = (count: number) =>
    applyDocumentOp(document, withNewIds(op, { revision: pool(count), control }));
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

const COMMENT_ANCHORS: ReadonlySet<string> = new Set([
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
]);

const leafPlan = ({ node, ancestors }: LeafSpan, author: string): LeafPlan => {
  if (ancestors.some(isRemovedRevisionNode)) return "untouched";
  if (COMMENT_ANCHORS.has(node.type)) return "anchor";
  const own = ancestors.some(
    (ancestor) =>
      ancestor.type === "insertion" &&
      isTrackedWrapper(ancestor) &&
      ancestor.info.author === author,
  );
  return own ? "direct" : "tracked";
};

type Segment = { plan: "direct" | "tracked"; from: Gap; to: Gap; records: boolean };

const laterGap = (left: Gap, right: Gap): Gap => (compareGaps(left, right) >= 0 ? left : right);
const earlierGap = (left: Gap, right: Gap): Gap => (compareGaps(left, right) <= 0 ? left : right);

/**
 * The stretches of a range to delete directly and with tracking, in document
 * order. A comment anchor ends a stretch; a tracked stretch runs across
 * deleted content, which it leaves alone, and one holding nothing else is
 * dropped.
 */
const segmentsOf = (spans: readonly LeafSpan[], from: Gap, to: Gap, author: string): Segment[] => {
  const out: Segment[] = [];
  let open: Segment | undefined;
  for (const span of spans) {
    const start = laterGap(span.before, from);
    const end = earlierGap(span.after, to);
    if (compareGaps(start, end) >= 0) continue;
    const plan = leafPlan(span, author);
    if (plan === "anchor") {
      open = undefined;
      continue;
    }
    const segmentPlan = plan === "untouched" ? "tracked" : plan;
    const records = plan !== "untouched";
    if (open !== undefined && open.plan === segmentPlan && compareGaps(open.to, start) === 0) {
      open.to = end;
      open.records ||= records;
      continue;
    }
    open = { plan: segmentPlan, from: start, to: end, records };
    out.push(open);
  }
  return out.filter(({ records }) => records);
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
  | SetParagraphPropsOp;

type TrackedPlanOptions = {
  document: Document;
  revision: RevisionStamp;
  newIds: NewIds | undefined;
};

/** Shared allocation and simulation for all pieces of one planned edit. */
export const createTrackedPlan = ({ document, revision, newIds }: TrackedPlanOptions) => {
  const used = usedIds(document, IDENTITY_SPACES.REVISION);
  const pool = [...new Set(newIds?.revision ?? [])].filter(
    (id) => !used.has(id) && id !== revision.id,
  );
  let taken = 0;
  let stampUsed = false;
  let current = document;
  const ops: DocumentOp[] = [];
  const take = (count: number) => {
    if (taken + count > pool.length) return undefined;
    const ids = pool.slice(taken, taken + count);
    taken += count;
    return ids;
  };
  const append = (input: PlannedReviewOp): Result<void, DocumentOpRefusal> => {
    let op: DocumentOp = input;
    if (input.revision !== undefined) {
      const stampId = stampUsed ? take(1)?.at(0) : revision.id;
      if (stampId === undefined) return outOfIds();
      op = { ...input, revision: { ...revision, id: stampId } };
      const demand = revisionIdDemand(current, op);
      if (demand.isErr()) return Result.err(demand.error);
      const ids = take(demand.value);
      if (ids === undefined) return outOfIds();
      const controls = usedIds(current, IDENTITY_SPACES.CONTROL);
      op = withNewIds(op, {
        revision: ids,
        control: (newIds?.control ?? []).filter((id) => !controls.has(id)),
      });
    }
    const applied = applyDocumentOp(current, op);
    if (applied.isErr()) return Result.err(applied.error);
    stampUsed ||= applied.value.revisions.length > 0;
    current = applied.value.document;
    ops.push(op);
    return Result.ok(undefined);
  };
  return { append, document: () => current, ops };
};

type TrackedPlan = ReturnType<typeof createTrackedPlan>;

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
    Result.err(new DocumentOpRefusal({ reason, message, opType: DOCUMENT_OP_TYPES.DELETE_RANGE }));
  if (from.story !== to.story)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A planned range must stay in one story.",
    );
  const body = storyBody(document, from.story);
  const paragraphs = storyParagraphs(body);
  const first = paragraphs.find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(from.blockId),
  );
  const last = paragraphs.find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(to.blockId),
  );
  if (first === undefined || last === undefined)
    return refuse(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND, "A range endpoint does not exist.");
  if (!sameBlockList(first.list, last.list))
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A planned range must stay in one block list.",
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
    const at = { story: from.story, blockId: paragraph.paraId ?? "", offset: 0 };
    const start = index === first.index ? from : { ...at, zeroWidthBefore: 0 };
    const end =
      index === last.index
        ? to
        : {
            ...at,
            offset: paragraphLength(paragraph),
            zeroWidthBefore: zeroWidthLeavesAt(paragraph.content, paragraphLength(paragraph))
              .length,
          };
    const check = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: start,
      to: end,
    });
    if (check.isErr()) return Result.err(check.error);
    if (first.index !== last.index) {
      if (index < last.index && paragraph.sectionProperties !== undefined)
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          "A planned range cannot cross a section boundary.",
        );
      if ((paragraph.propertyChanges?.length ?? 0) > 0)
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          "Cross-paragraph planning cannot replace an existing property review.",
        );
    }
    const content = paragraph.content;
    const segments = segmentsOf(
      leafSpans(content),
      {
        offset: start.offset,
        zeroWidthBefore: start.zeroWidthBefore ?? zeroWidthLeavesAt(content, start.offset).length,
      },
      { offset: end.offset, zeroWidthBefore: end.zeroWidthBefore ?? 0 },
      revision.author,
    );
    ranges.push({ paragraph, at, segments });
  }
  // Direct simulation gives the surviving paragraph's authored properties,
  // independent of tracked-deleted content that still occupies offsets.
  let direct = document;
  for (const { at, segments } of ranges.toReversed()) {
    for (const segment of segments.toReversed()) {
      const op = {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: positionOf(at, segment.from),
        to: positionOf(at, segment.to),
      } satisfies DeleteRangeOp;
      const appended = plan.append(segment.plan === "tracked" ? { ...op, revision } : op);
      if (appended.isErr()) return appended;
      const simulated = applyDocumentOp(direct, op);
      if (simulated.isErr()) return Result.err(simulated.error);
      direct = simulated.value.document;
    }
  }
  if (first.index === last.index) return Result.ok(undefined);
  for (let index = ranges.length - 2; index >= 0; index -= 1) {
    const leading = ranges[index];
    const following = ranges[index + 1];
    if (leading === undefined || following === undefined)
      return refuse(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE, "A planned join has no endpoint.");
    const op = {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: from.story,
      blockId: leading.at.blockId,
      nextBlockId: following.at.blockId,
      revision,
    } satisfies JoinBlocksOp;
    const appended = plan.append(op);
    if (appended.isErr()) return appended;
    const simulated = applyDocumentOp(direct, {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: from.story,
      blockId: leading.at.blockId,
      nextBlockId: to.blockId,
    });
    if (simulated.isErr()) return Result.err(simulated.error);
    direct = simulated.value.document;
  }
  const survivor = storyParagraphs(storyBody(direct, from.story)).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(to.blockId),
  );
  const trackedSurvivor = storyParagraphs(storyBody(plan.document(), from.story)).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(to.blockId),
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
  const plan = createTrackedPlan({ document, revision: options.revision, newIds: options.newIds });
  const appended = appendTrackedDeletion({ document, options, plan });
  return appended.isErr() ? Result.err(appended.error) : Result.ok(plan.ops);
};
