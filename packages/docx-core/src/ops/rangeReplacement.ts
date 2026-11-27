/** Tracked replacement fragments share deletion's validation and physical-id pool. */
import { panic, Result } from "better-result";

import type { Document, DocumentBody, Paragraph } from "../model/document";
import { storyBody, storyParagraphs } from "./blocks";
import { newlyExhaustedCommentIds } from "./comments";
import { identityKeysIn, idKey, slotKey } from "./ids";
import { gapAfterInserted } from "./inline";
import {
  asParagraphContent,
  childNodes,
  compareGaps,
  defaultInsertionGap,
  type Gap,
  type InlineNode,
  isParagraphContent,
  leafSpans,
  rebuildNode,
  sameOwnFields,
  spanningRecords,
  zeroWidthLeavesAt,
} from "./leaves";
import { itemBefore } from "./listItems";
import { isClosingMarker, paragraphLength } from "./offsets";
import { findLastStoryParagraph, findStoryParagraph } from "./paragraphLocator";
import {
  appendTrackedDeletion,
  createTrackedPlan,
  type PlanTrackedDeletionOptions,
  selectedParagraphRuns,
} from "./plan";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { isRemovedRevisionNode } from "./review";
import { isOwnRecord } from "./revisionOwner";
import { identitySlots } from "./slots";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InlineSlice,
  type OpStory,
  type ReplaceInlineOp,
  type RevisionStamp,
  SPLIT_HALVES,
  type TextPosition,
} from "./types";

type RetainedEnclosing = NonNullable<
  NonNullable<Extract<InlineNode, { type: "deletion" }>["resolutionJoins"]>
>["retainedEnclosing"];

const rangeParagraph = (body: DocumentBody, blockId: string) =>
  blockId === ""
    ? storyParagraphs(body).find(
        ({ paragraph }) => (paragraph.paraId ?? "") === "",
      )
    : findStoryParagraph(body, blockId);

type RangeEndAfterDeletionOptions = {
  before: Document;
  after: Document;
  from: TextPosition;
  to: TextPosition;
};

/**
 * Where a text replacement leaves the comment anchors it reaches. The direct
 * edit and the accepted suggestion give the same answer, so a comment's fate
 * never depends on the mode:
 *
 * - A comment marker (range start, range end or reference) inside the
 *   replaced stretch is not deleted with the words. It stays, in its order,
 *   and the new words follow it (this gap). A comment straddling an edge
 *   keeps its outer marker: one ending inside now ends where the new words
 *   start, one starting inside now covers the new words and the rest of its
 *   own.
 * - A ranged comment whose every word went has nothing left to cover, and
 *   its thread goes, replies included, directly and once the suggestion is
 *   accepted (`newlyExhaustedCommentIds`). A point comment, or a range that
 *   covered nothing before, keeps its thread and stands where the new words
 *   start.
 * - Tracked, the struck words keep every marker among them
 *   (`replacementDeletionSegments` ends a stretch at each), so rejecting
 *   restores each anchor exactly and accepting leaves the direct result.
 * - A replacement across paragraphs joins them: the markers of every
 *   paragraph it reaches gather, in order, before the new words.
 *
 * The gap just after a range deletion, where what replaces the range goes:
 * after the struck content in a redline, and after whatever a deletion keeps
 * in place, so that accepting a tracked replacement gives the direct one.
 * A range across containers is replaced in its first paragraph run, at that
 * run's end. Content after the range's end is untouched by the deletion, so
 * the gap is found by its distance from the end of the paragraph now holding
 * that content; joins may have moved it into an earlier paragraph of the run.
 * `after` may also hold content inserted at the gap; the gap then follows it.
 */
export const rangeEndAfterDeletion = ({
  before,
  after,
  from,
  to,
}: RangeEndAfterDeletionOptions): TextPosition => {
  const runs = selectedParagraphRuns(before, from, to);
  if (runs.isErr())
    panic("A successful range deletion must have its input paragraphs.");
  const run = runs.value.at(0) ?? [];
  const source = run.at(-1)?.paragraph;
  if (source === undefined)
    panic("A successful range deletion must select a paragraph.");
  const sourceLength = paragraphLength(source);
  // An empty range is a caret: it stands where an insertion defaults to.
  const caret =
    idKey(from.blockId) === idKey(to.blockId) &&
    from.offset === to.offset &&
    from.zeroWidthBefore === to.zeroWidthBefore;
  const end =
    idKey(source.paraId ?? "") === idKey(to.blockId)
      ? {
          offset: to.offset,
          zeroWidthBefore:
            to.zeroWidthBefore ??
            (caret
              ? defaultInsertionGap(source.content, to.offset).zeroWidthBefore
              : 0),
        }
      : {
          offset: sourceLength,
          zeroWidthBefore: zeroWidthLeavesAt(source.content, sourceLength)
            .length,
        };
  const body = storyBody(after, from.story);
  let holder: Paragraph | undefined;
  let target: Paragraph | undefined;
  let checkedUnnamed = false;
  for (let index = run.length - 1; index >= 0; index -= 1) {
    const candidate = run.at(index)?.paragraph;
    if (candidate === undefined) continue;
    const key = idKey(candidate.paraId ?? "");
    // The old story map kept the last physical occurrence, including unnamed
    // paragraphs and duplicate IDs. The locator preserves that order for IDs.
    if (key === "") {
      if (checkedUnnamed) continue;
      checkedUnnamed = true;
      target = storyParagraphs(body).findLast(
        ({ paragraph }) => (paragraph.paraId ?? "") === "",
      )?.paragraph;
    } else {
      target = findLastStoryParagraph(body, key)?.paragraph;
    }
    if (target !== undefined) {
      holder = candidate;
      break;
    }
  }
  if (holder === undefined || target === undefined)
    panic("A successful range deletion must retain a selected paragraph.");
  const offset = paragraphLength(target) - (sourceLength - end.offset);
  const following =
    zeroWidthLeavesAt(source.content, end.offset).length - end.zeroWidthBefore;
  const zeroWidthBefore =
    zeroWidthLeavesAt(target.content, offset).length - following;
  const gap = { offset, zeroWidthBefore };
  return {
    story: from.story,
    blockId: holder.paraId ?? to.blockId,
    ...(caret ? gap : pastEndedComments(target.content, gap)),
  };
};

/**
 * Where what replaces a range goes among the zero-width leaves at the gap:
 * past the ends of comments the range reached, so the new words never extend
 * a comment whose end the range took, and past the ends of other ranges
 * standing before such a comment end, including references interleaved with
 * those ends. A comment starting at the gap keeps the new words, as it does when
 * a tracked replacement is accepted. The new words continue the containers
 * the range's text stood in across the ends they pass (see `continuation`).
 */
export const pastEndedComments = (
  content: readonly InlineNode[],
  gap: Gap,
): Gap => {
  const spans = leafSpans(content);
  const first = spans.findIndex(({ before }) => compareGaps(before, gap) === 0);
  if (first === -1) return gap;
  let passed = gap;
  for (let index = first; index < spans.length; index += 1) {
    const span = spans[index];
    if (span === undefined || !isParagraphContent(span.node)) break;
    const { node } = span;
    if (node.type === "commentRangeEnd") passed = span.after;
    else if (node.type !== "commentReference" && !isClosingMarker(node)) break;
  }
  return passed;
};

/**
 * A range replacement fragment: newly identified paragraphs before its last
 * paragraph, and the inline tail inserted into the retained original paragraph.
 * Each new paragraph's own fields apply to it, including when it takes the
 * source prefix. The last fragment keeps the original survivor's identity.
 * Closed inline slices are required. Imported review and content-control ids
 * must already be distinct from destination identities; section-bearing
 * paragraphs require an explicit package import policy.
 */
export type PlanTrackedReplaceOptions = PlanTrackedDeletionOptions & {
  /** How the deletion carrying the edit's own id is recorded, when not as `revision`. */
  deletionRevision?: RevisionStamp;
  sourceContainerPolicy: "join" | "separate";
  seamPolicy?: Extract<DocumentOp, { type: "insertContent" }>["seamPolicy"];
  replacement: {
    paragraphs: readonly Paragraph[];
    tail: InlineSlice;
  };
};

type PlanTrackedReplacementOptions = PlanTrackedReplaceOptions & {
  cuts?: readonly TextPosition[];
  retainedControlInsertion?: { at: TextPosition; slice: InlineSlice };
  retainControls?: {
    ranges: readonly { start: TextPosition; end: TextPosition }[];
    moveInside: (options: {
      document: Document;
      start: TextPosition;
      operations: readonly DocumentOp[];
    }) => ReplaceInlineOp | undefined;
  };
};

/**
 * Plan the replacement with its deletion cut at the supplied source gaps.
 * Every piece draws revision ids from the same pool.
 */
export const planTrackedReplacement = (
  document: Document,
  options: PlanTrackedReplacementOptions,
): Result<DocumentOp[], DocumentOpRefusal> => {
  const { replacement, from, revision, newIds, cuts = [] } = options;
  const refuse = (message: string) =>
    Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        message,
        opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      }),
    );
  if (replacement.tail.openStart !== 0 || replacement.tail.openEnd !== 0) {
    return refuse("Tracked replacement requires a closed tail slice.");
  }
  if (
    replacement.paragraphs.some(
      ({ sectionProperties }) => sectionProperties !== undefined,
    )
  ) {
    return refuse(
      "Tracked replacement cannot import section-bearing paragraphs.",
    );
  }
  const plan = createTrackedPlan({
    document,
    revision,
    newIds,
    ...(options.deletionRevision === undefined
      ? {}
      : { deletionRevision: options.deletionRevision }),
  });
  const edges = [from, ...cuts, options.to];
  const sameGap = (left: TextPosition, right: TextPosition) =>
    left.offset === right.offset &&
    (left.zeroWidthBefore ?? 0) === (right.zeroWidthBefore ?? 0);
  const pieces = [];
  for (let index = edges.length - 1; index > 0; index -= 1) {
    const start = itemBefore(edges, index);
    const end = edges.at(index);
    if (start === undefined || end === undefined)
      panic("A replacement deletion has both endpoints.");
    const retainsControl =
      options.retainControls?.ranges.some(
        (range) => sameGap(start, range.start) && sameGap(end, range.end),
      ) === true;
    pieces.push({ start, end, retainsControl });
  }
  // Move the control's deletion inside it before the neighbouring pieces
  // can join its record. These pieces only stage tracked deletions, so the
  // following source positions are unchanged.
  for (const { start, end, retainsControl } of pieces.toSorted(
    (left, right) => Number(right.retainsControl) - Number(left.retainsControl),
  )) {
    const operationStart = plan.ops.length;
    const deleted = appendTrackedDeletion({
      document: plan.document(),
      options: { ...options, from: start, to: end },
      plan,
    });
    if (deleted.isErr()) return Result.err(deleted.error);
    if (!retainsControl) continue;
    const move = options.retainControls?.moveInside({
      document: plan.document(),
      start,
      operations: plan.ops.slice(operationStart),
    });
    if (move !== undefined) {
      const moved = plan.append(move);
      if (moved.isErr()) return Result.err(moved.error);
    }
  }
  if (
    options.sourceContainerPolicy === "separate" &&
    (replacement.tail.content.length > 0 ||
      replacement.paragraphs.some((paragraph) => paragraph.content.length > 0))
  ) {
    const deletionIds = new Set(
      plan.ops.flatMap((op) =>
        op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && op.revision !== undefined
          ? [op.revision.id]
          : [],
      ),
    );
    const deletionBlocks = new Set(
      plan.ops.flatMap((op) => {
        if (op.type === DOCUMENT_OP_TYPES.DELETE_RANGE)
          return [idKey(op.from.blockId)];
        if (op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS)
          return [idKey(op.nextBlockId)];
        return [];
      }),
    );
    const breakRemovedSeam = (node: InlineNode): InlineNode => {
      let next = node;
      if (
        node.type === "deletion" &&
        deletionIds.has(node.info.id) &&
        node.resolutionJoins !== undefined &&
        node.resolutionJoins.remove > 0
      ) {
        next = Object.assign({}, node, {
          resolutionJoins: Object.assign({}, node.resolutionJoins, {
            remove: 0,
          }),
        });
      }
      const children = childNodes(next);
      if (children === undefined) return next;
      const changed = children.map(breakRemovedSeam);
      return changed.every((child, index) => child === children[index])
        ? next
        : rebuildNode(next, changed);
    };
    // The new authored payload separates the source's two cut pieces.
    // Removing this deletion must retain its identity transfers without
    // joining the payload to the retained suffix.
    for (const { paragraph } of storyParagraphs(
      storyBody(plan.document(), from.story),
    )) {
      if (!deletionBlocks.has(idKey(paragraph.paraId ?? ""))) continue;
      const content = paragraph.content.map(breakRemovedSeam);
      if (content.every((node, index) => node === paragraph.content[index]))
        continue;
      const patched = plan.append({
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story: from.story,
        blockId: paragraph.paraId ?? "",
        expected: paragraph.content,
        content: asParagraphContent(content),
      });
      if (patched.isErr()) return Result.err(patched.error);
    }
  }
  // Tracking leaves selected content in place. The replacement follows it, as
  // the struck words come first in a redline; splitting moves the old mark to
  // the original half.
  let at = rangeEndAfterDeletion({
    before: document,
    after: plan.document(),
    from,
    to: options.to,
  });
  for (const paragraph of replacement.paragraphs) {
    const sourceBlockId = at.blockId;
    const source = rangeParagraph(
      storyBody(plan.document(), from.story),
      sourceBlockId,
    );
    if (source === undefined)
      return refuse("The replacement's source paragraph does not exist.");
    const gap =
      at.zeroWidthBefore === undefined
        ? defaultInsertionGap(source.paragraph.content, at.offset)
        : { offset: at.offset, zeroWidthBefore: at.zeroWidthBefore };
    const end = gapAfterInserted(gap, paragraph.content);
    if (paragraph.content.length > 0) {
      const inserted = plan.append({
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at,
        slice: { content: paragraph.content, openStart: 0, openEnd: 0 },
        ...(options.seamPolicy === undefined
          ? {}
          : { seamPolicy: options.seamPolicy }),
        revision,
      });
      if (inserted.isErr()) return Result.err(inserted.error);
    }
    const {
      type: _type,
      paraId,
      content: _content,
      sectionProperties: _section,
      pPrMark: _mark,
      propertyChanges: _changes,
      ...newParagraph
    } = paragraph;
    const split = plan.append({
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: { ...at, ...end },
      newBlockId: paraId ?? "",
      newHalf: SPLIT_HALVES.FIRST,
      newParagraph,
      revision,
    });
    if (split.isErr()) return Result.err(split.error);
    at = {
      story: from.story,
      blockId: at.blockId,
      offset: 0,
      zeroWidthBefore: 0,
    };
  }
  // New paragraphs part the pieces of a record the deletion cut at the
  // range's end. The first fragment records that cut as an insertion made
  // there would, so rejecting the changes in any order rejoins the pieces.
  const leadingFragment = replacement.paragraphs.at(0);
  const toParagraph =
    leadingFragment === undefined || leadingFragment.content.length === 0
      ? undefined
      : rangeParagraph(storyBody(document, from.story), options.to.blockId)
          ?.paragraph;
  const cutDepth =
    toParagraph === undefined
      ? 0
      : spanningRecords(
          toParagraph.content,
          [
            {
              offset: options.to.offset,
              zeroWidthBefore: options.to.zeroWidthBefore ?? 0,
            },
          ],
          0,
          1,
        ).length;
  if (
    cutDepth > 0 &&
    leadingFragment !== undefined &&
    leadingFragment.content.length > 0
  ) {
    const first = rangeParagraph(
      storyBody(plan.document(), from.story),
      leadingFragment.paraId ?? "",
    )?.paragraph;
    const payload = first?.content.at(-1);
    if (
      first !== undefined &&
      payload?.type === "insertion" &&
      isOwnRecord(payload.info, revision) &&
      (payload.resolutionJoins?.remove ?? 0) < cutDepth
    ) {
      const recorded = plan.append({
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story: from.story,
        blockId: first.paraId ?? "",
        expected: first.content,
        content: asParagraphContent([
          ...first.content.slice(0, -1),
          Object.assign({}, payload, {
            resolutionJoins: Object.assign(
              { before: 0, after: 0, remove: 0 },
              payload.resolutionJoins,
              { remove: cutDepth },
            ),
          }),
        ]),
      });
      if (recorded.isErr()) return Result.err(recorded.error);
    }
  }
  if (replacement.tail.content.length > 0) {
    const inserted = plan.append({
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: options.retainedControlInsertion?.at ?? at,
      slice: options.retainedControlInsertion?.slice ?? replacement.tail,
      ...(options.seamPolicy === undefined
        ? {}
        : { seamPolicy: options.seamPolicy }),
      revision,
    });
    if (inserted.isErr()) return Result.err(inserted.error);
  }
  // Retracting one's own pending insertion physically removes text. Any
  // range it exhausts goes with that text, even before review is resolved.
  if (
    plan.ops.some(
      (op) =>
        op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && op.revision === undefined,
    )
  ) {
    for (const id of newlyExhaustedCommentIds(document, plan.document()))
      plan.ops.push({
        type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
        id,
        scope: "thread",
      });
  }
  return Result.ok(plan.ops);
};

/** Plan a same-list range deletion followed by tracked replacement insertions. */
export const planTrackedReplace = (
  document: Document,
  options: PlanTrackedReplaceOptions,
): Result<DocumentOp[], DocumentOpRefusal> =>
  planTrackedReplacement(document, options);

type CutPieceHandoffsOptions = {
  before: Document;
  after: Document;
  story: OpStory;
};

/**
 * Handoffs for the records an edit cut whose first piece holds only content
 * on its way out. Accepting that removal empties the first piece, which then
 * goes; the piece cut off after it, wherever the edit left it in the story,
 * takes back the record's identities, as the same edit made directly leaves
 * the record whole. Each removal in the first piece carries the handoff.
 */
export const cutPieceHandoffs = ({
  before,
  after,
  story,
}: CutPieceHandoffsOptions): ReplaceInlineOp[] => {
  const present = new Set(identityKeysIn(storyBody(before, story).content));
  const paragraphs = storyParagraphs(storyBody(after, story)).map(
    ({ paragraph }) => paragraph,
  );
  type Located = { node: InlineNode; paragraph: number };
  const order: Located[] = [];
  const visit = (nodes: readonly InlineNode[], paragraph: number) => {
    for (const node of nodes) {
      order.push({ node, paragraph });
      visit(childNodes(node) ?? [], paragraph);
    }
  };
  for (const [index, paragraph] of paragraphs.entries())
    visit(paragraph.content, index);
  const handoffs = new Map<InlineNode, NonNullable<RetainedEnclosing>>();
  for (const [index, { node: first }] of order.entries()) {
    const children = childNodes(first);
    const source = identitySlots(first);
    if (
      children === undefined ||
      children.length === 0 ||
      !children.every(isRemovedRevisionNode) ||
      source.length === 0 ||
      !source.every((slot) => present.has(slotKey(slot)))
    )
      continue;
    // The next record alike in fields is the piece cut off this one when it
    // is new; when it is another record of the source, this one was not cut,
    // and a new piece further on was cut off that other record.
    const next = order
      .slice(index + 1)
      .find(
        ({ node }) =>
          sameOwnFields(node, first) &&
          identitySlots(node).length === source.length,
      )?.node;
    const piece =
      next !== undefined &&
      identitySlots(next).every((slot) => !present.has(slotKey(slot)))
        ? next
        : undefined;
    if (piece === undefined) continue;
    const target = identitySlots(piece);
    for (const removal of children)
      handoffs.set(removal, [
        ...(handoffs.get(removal) ?? []),
        { levels: 1, source, target },
      ]);
  }
  if (handoffs.size === 0) return [];
  const carry = (node: InlineNode): InlineNode => {
    const entries = handoffs.get(node);
    let next = node;
    if (
      entries !== undefined &&
      (node.type === "deletion" || node.type === "moveFrom")
    ) {
      const joins = node.resolutionJoins ?? { before: 0, after: 0, remove: 0 };
      next = Object.assign({}, node, {
        resolutionJoins: Object.assign({}, joins, {
          retainedEnclosing: [...(joins.retainedEnclosing ?? []), ...entries],
        }),
      });
    }
    const children = childNodes(next);
    if (children === undefined) return next;
    const changed = children.map(carry);
    return changed.every((child, index) => child === children[index])
      ? next
      : rebuildNode(next, changed);
  };
  return paragraphs.flatMap((paragraph) => {
    const content = paragraph.content.map(carry);
    return content.every((node, index) => node === paragraph.content[index])
      ? []
      : [
          {
            type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
            story,
            blockId: paragraph.paraId ?? "",
            expected: paragraph.content,
            content: asParagraphContent(content),
          },
        ];
  });
};
