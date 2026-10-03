import { findStoryBody, sameStory, storyBody } from "./stories";
/**
 * Pure planning around tracked operations: how many new revision ids an
 * operation takes, and the operations a tracked deletion of a range is made
 * of once the author's own insertions and comment anchors are accounted for.
 */

import { Result, panic } from "better-result";

import { type Document, MAX_REVISION_ID } from "../model/document";
import { applyDocumentOp, stampOf, type AppliedDocumentOp } from "./apply";
import { blockListAt, sameBlockList, storyParagraphs, type ParagraphLocation } from "./blocks";
import { structurallyEqual } from "./equality";
import {
  IDENTITY_SPACES,
  type IdentitySpace,
  identityKeysIn,
  idKey,
  packageIdentityKeys,
} from "./ids";
import { paragraphLength } from "./offsets";
import {
  compareGaps,
  isCommentAnchor,
  type Gap,
  type LeafSpan,
  leafSpans,
  zeroWidthLeavesAt,
} from "./leaves";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  isAddedRevision,
  isRemovedRevisionNode,
  isTrackedWrapper,
  paragraphPropertiesOf,
} from "./review";
import {
  DOCUMENT_OP_TYPES,
  SECTION_BOUNDARY_POLICIES,
  type DeleteRangeOp,
  type InsertContentOp,
  type JoinBlocksOp,
  type SetParagraphPropsOp,
  type SetRunPropsOp,
  type SplitBlockOp,
  EMPTY_PROPERTY_SETS,
  type DocumentOp,
  type NewIds,
  type OpStory,
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
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
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
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
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
    newIds.revision = revision.filter((id) => keep(IDENTITY_SPACES.REVISION, id));
  if (control !== undefined)
    newIds.control = control.filter((id) => keep(IDENTITY_SPACES.CONTROL, id));
  return withNewIds(op, newIds);
};

type TrimAppliedNewIdsOptions = {
  op: DocumentOp;
  applied: AppliedDocumentOp;
  story: OpStory;
};

/** Keep only supplied ids that the successful operation put in changed paragraphs. */
export const trimAppliedNewIds = ({ op, applied, story }: TrimAppliedNewIdsOptions): DocumentOp => {
  if (!("newIds" in op) || op.newIds === undefined) return op;
  const touched = new Set([...applied.touched.modified, ...applied.touched.inserted].map(idKey));
  const paragraphs = storyParagraphs(storyBody(applied.document, story))
    .map(({ paragraph }) => paragraph)
    .filter(({ paraId }) => paraId !== undefined && touched.has(idKey(paraId)));
  const identities = new Set(identityKeysIn(paragraphs));
  return filterNewIds(op, (space, id) => identities.has(`${space}:${id}`));
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

const leafPlan = ({ node, ancestors }: LeafSpan, author: string): LeafPlan => {
  if (ancestors.some(isRemovedRevisionNode)) return "untouched";
  if (isCommentAnchor(node)) return "anchor";
  const own = ancestors.some(
    (ancestor) =>
      isAddedRevision(ancestor) && isTrackedWrapper(ancestor) && ancestor.info.author === author,
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
type ReplacementDeletionSegmentsOptions = {
  spans: readonly LeafSpan[];
  from: Gap;
  to: Gap;
  mode: { type: "editing" } | { type: "suggesting"; author: string };
};

export const replacementDeletionSegments = ({
  spans,
  from,
  to,
  mode,
}: ReplacementDeletionSegmentsOptions): Segment[] => {
  const out: Segment[] = [];
  let open: Segment | undefined;
  for (const span of spans) {
    const start = laterGap(span.before, from);
    const end = earlierGap(span.after, to);
    if (compareGaps(start, end) >= 0) continue;
    let plan: LeafPlan = isCommentAnchor(span.node) ? "anchor" : "direct";
    if (mode.type === "suggesting") plan = leafPlan(span, mode.author);
    else if (span.ancestors.some(isRemovedRevisionNode)) plan = "untouched";
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
  | SetParagraphPropsOp
  | SetRunPropsOp;

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
    if (input.revision !== undefined) {
      // Source cuts consume the caller's pool in the same order as editing.
      // Additional mode-owned stamps use its other end and cannot displace them.
      let stampId = revision.id;
      if (stampUsed) {
        if (remaining <= taken) return outOfIds();
        remaining -= 1;
        stampId = pool.at(remaining) ?? panic("A reserved stamp has an allocated pool entry.");
      }
      op = { ...input, revision: { ...revision, id: stampId } };
    }
    const demand = revisionIdDemand(current, op);
    if (demand.isErr()) return Result.err(demand.error);
    const ids = take(demand.value);
    if (ids === undefined) return outOfIds();
    const controls = usedIds(current, IDENTITY_SPACES.CONTROL);
    op = withNewIds(op, {
      revision: ids,
      control: (newIds?.control ?? []).filter((id) => !controls.has(id)),
    });
    const applied = applyDocumentOp(current, op);
    if (applied.isErr()) return Result.err(applied.error);
    stampUsed ||= input.revision !== undefined && applied.value.revisions.length > 0;
    current = applied.value.document;
    const story = (() => {
      switch (input.type) {
        case DOCUMENT_OP_TYPES.DELETE_RANGE:
        case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
          return input.from.story;
        case DOCUMENT_OP_TYPES.INSERT_CONTENT:
        case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
          return input.at.story;
        case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
        case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
          return input.story;
        default: {
          const unreachable: never = input;
          return unreachable;
        }
      }
    })();
    ops.push(trimAppliedNewIds({ op, applied: applied.value, story }));
    return Result.ok(undefined);
  };
  return { append, document: () => current, ops };
};

type TrackedPlan = ReturnType<typeof createTrackedPlan>;

/** A text range crosses containers without removing their geometry or boundary marks. */
export const selectedParagraphRuns = (
  document: Document,
  from: TextPosition,
  to: TextPosition,
): Result<ParagraphLocation[][], DocumentOpRefusal> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: DOCUMENT_OP_TYPES.DELETE_RANGE }));
  if (!sameStory(from.story, to.story))
    return refuse(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE, "A text range stays in one story.");
  const locations = storyParagraphs(storyBody(document, from.story));
  const first = locations.findIndex(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(from.blockId),
  );
  const last = locations.findIndex(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(to.blockId),
  );
  if (first < 0 || last < 0)
    return refuse(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND, "A range endpoint does not exist.");
  if (first > last)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
      "Range endpoints must be in document order.",
    );
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
    Result.err(new DocumentOpRefusal({ reason, message, opType: DOCUMENT_OP_TYPES.DELETE_RANGE }));
  const partitioned = selectedParagraphRuns(document, from, to);
  if (partitioned.isErr()) return Result.err(partitioned.error);
  if (partitioned.value.length > 1) {
    for (const run of partitioned.value.toReversed()) {
      const first = run.at(0)?.paragraph;
      const last = run.at(-1)?.paragraph;
      if (!first || !last) panic("A selected paragraph run must contain its endpoints.");
      const start =
        idKey(first.paraId ?? "") === idKey(from.blockId)
          ? from
          : { story: from.story, blockId: first.paraId ?? "", offset: 0, zeroWidthBefore: 0 };
      const end =
        idKey(last.paraId ?? "") === idKey(to.blockId)
          ? to
          : {
              story: to.story,
              blockId: last.paraId ?? "",
              offset: paragraphLength(last),
              zeroWidthBefore: zeroWidthLeavesAt(last.content, paragraphLength(last)).length,
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
    return refuse(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND, "The story does not exist.");
  const paragraphs = storyParagraphs(body);
  const first = paragraphs.find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(from.blockId),
  );
  const last = paragraphs.find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(to.blockId),
  );
  if (first === undefined || last === undefined)
    return refuse(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND, "A range endpoint does not exist.");
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
    const check = revisionIdDemand(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: start,
      to: end,
    });
    if (check.isErr()) return Result.err(check.error);
    const content = paragraph.content;
    const segments = replacementDeletionSegments({
      spans: leafSpans(content),
      from: {
        offset: start.offset,
        zeroWidthBefore: start.zeroWidthBefore ?? zeroWidthLeavesAt(content, start.offset).length,
      },
      to: { offset: end.offset, zeroWidthBefore: end.zeroWidthBefore ?? 0 },
      mode: { type: "suggesting", author: revision.author },
    });
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
      const allocated = plan.ops.at(-1);
      const simulated = applyDocumentOp(direct, {
        ...op,
        ...(allocated?.type === DOCUMENT_OP_TYPES.DELETE_RANGE ? { newIds: allocated.newIds } : {}),
      });
      if (simulated.isErr()) return Result.err(simulated.error);
      direct = simulated.value.document;
    }
  }
  if (first.index === last.index) return Result.ok(undefined);
  for (let index = ranges.length - 2; index >= 0; index -= 1) {
    const leading = ranges[index];
    if (leading === undefined)
      return refuse(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE, "A planned join has no endpoint.");
    const currentBody = storyBody(plan.document(), from.story);
    const currentLeading = storyParagraphs(currentBody).find(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(leading.at.blockId),
    );
    if (currentLeading === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        "A planned join lost its leading paragraph.",
      );
    const mark = currentLeading.paragraph.pPrMark;
    if (mark?.kind === "del" || mark?.kind === "moveFrom") continue;
    const ownAddedMark =
      (mark?.kind === "ins" || mark?.kind === "moveTo") && mark.info.author === revision.author;
    const next = blockListAt(currentBody.content, currentLeading.list).at(currentLeading.index + 1);
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
    const appended = plan.append(ownAddedMark ? join : { ...join, revision });
    if (appended.isErr()) return appended;
    const directBody = storyBody(direct, from.story);
    const directLeading = storyParagraphs(directBody).find(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(leading.at.blockId),
    );
    if (directLeading === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        "A simulated join lost its leading paragraph.",
      );
    const directNext = blockListAt(directBody.content, directLeading.list).at(
      directLeading.index + 1,
    );
    if (directNext?.type !== "paragraph")
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
        "A simulated join has no following paragraph.",
      );
    const simulated = applyDocumentOp(direct, { ...join, nextBlockId: directNext.paraId ?? "" });
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
