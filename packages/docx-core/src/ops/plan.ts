/**
 * Pure planning around tracked operations: how many new revision ids an
 * operation takes, and the operations a tracked deletion of a range is made
 * of once the author's own insertions and comment anchors are accounted for.
 */

import { Result } from "better-result";

import { type Document, MAX_REVISION_ID } from "../model/document";
import { applyDocumentOp, stampOf } from "./apply";
import { storyBody, storyParagraphs } from "./blocks";
import {
  IDENTITY_SPACES,
  type IdentitySpace,
  identityKeysIn,
  idKey,
  packageIdentityKeys,
} from "./ids";
import { compareGaps, type Gap, type LeafSpan, leafSpans, zeroWidthLeavesAt } from "./leaves";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { isRemovedRevisionNode, isTrackedWrapper } from "./review";
import {
  DOCUMENT_OP_TYPES,
  type DeleteRangeOp,
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
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return { ...op, newIds };
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
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

/**
 * The operations a tracked deletion of a range is made of, to apply in
 * order as one batch: a direct `deleteRange` for each stretch of the author's
 * own insertions (retracting a suggestion removes it), and a tracked one for
 * each other stretch, split around comment boundaries and references, which
 * no tracked change can hold and which therefore stay. Stretches go from the
 * last to the first, so each position still names what it named in the
 * input. The positions default their `zeroWidthBefore` as `deleteRange` does.
 *
 * The first tracked deletion carries the stamp's id. Each later one takes the
 * next id of `newIds.revision` the package does not use for its stamp, then
 * the ones it needs for the records it cuts.
 */
export const planTrackedDeletion = (
  document: Document,
  options: PlanTrackedDeletionOptions,
): Result<DocumentOp[], DocumentOpRefusal> => {
  const { from, to, revision, newIds } = options;
  // The direct deletion of the range validates its positions.
  const check = applyDocumentOp(document, { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from, to });
  if (check.isErr()) {
    return Result.err(check.error);
  }
  const location = storyParagraphs(storyBody(document, from.story)).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(from.blockId),
  );
  if (location === undefined) {
    return Result.ok([]);
  }
  const { content } = location.paragraph;
  const start: Gap = {
    offset: from.offset,
    zeroWidthBefore: from.zeroWidthBefore ?? zeroWidthLeavesAt(content, from.offset).length,
  };
  const end: Gap = { offset: to.offset, zeroWidthBefore: to.zeroWidthBefore ?? 0 };
  const segments = segmentsOf(leafSpans(content), start, end, revision.author);

  const used = usedIds(document, IDENTITY_SPACES.REVISION);
  const pool = (newIds?.revision ?? []).filter((id) => !used.has(id) && id !== revision.id);
  let taken = 0;
  const take = (count: number): number[] | undefined => {
    if (taken + count > pool.length) return undefined;
    const ids = pool.slice(taken, taken + count);
    taken += count;
    return ids;
  };

  const ops: DocumentOp[] = [];
  let current = document;
  let stampUsed = false;
  for (const segment of segments.toReversed()) {
    let op: DeleteRangeOp = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: positionOf(from, segment.from),
      to: positionOf(from, segment.to),
    };
    if (segment.plan === "tracked") {
      const stampId = stampUsed ? take(1)?.[0] : revision.id;
      if (stampId === undefined) return outOfIds();
      stampUsed = true;
      op = { ...op, revision: { ...revision, id: stampId } };
      const demand = revisionIdDemand(current, op);
      if (demand.isErr()) return Result.err(demand.error);
      const ids = take(demand.value);
      if (ids === undefined) return outOfIds();
      if (ids.length > 0) op = { ...op, newIds: { revision: ids } };
    }
    const applied = applyDocumentOp(current, op);
    if (applied.isErr()) return Result.err(applied.error);
    current = applied.value.document;
    ops.push(op);
  }
  return Result.ok(ops);
};
