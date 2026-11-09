/** Tracked replacement fragments share deletion's validation and physical-id pool. */
import { panic, Result } from "better-result";

import type { Document, Paragraph } from "../model/document";
import { storyBody, storyParagraphs } from "./blocks";
import { identityKeysIn, idKey, slotKey } from "./ids";
import { gapAfterInserted } from "./inline";
import {
  asParagraphContent,
  childNodes,
  defaultInsertionGap,
  type InlineNode,
  rebuildNode,
  sameOwnFields,
  spanningRecords,
  zeroWidthLeavesAt,
} from "./leaves";
import { paragraphLength } from "./offsets";
import {
  appendTrackedDeletion,
  createTrackedPlan,
  type PlanTrackedDeletionOptions,
  selectedParagraphRuns,
} from "./plan";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { isRemovedRevisionNode } from "./review";
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

type RangeEndAfterDeletionOptions = {
  before: Document;
  after: Document;
  from: TextPosition;
  to: TextPosition;
};

/**
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
  const current = new Map(
    storyParagraphs(storyBody(after, from.story)).map(({ paragraph }) => [
      idKey(paragraph.paraId ?? ""),
      paragraph,
    ]),
  );
  const holder = run.findLast(({ paragraph }) =>
    current.has(idKey(paragraph.paraId ?? "")),
  )?.paragraph;
  const target =
    holder === undefined ? undefined : current.get(idKey(holder.paraId ?? ""));
  if (holder === undefined || target === undefined)
    panic("A successful range deletion must retain a selected paragraph.");
  const offset = paragraphLength(target) - (sourceLength - end.offset);
  const following =
    zeroWidthLeavesAt(source.content, end.offset).length - end.zeroWidthBefore;
  return {
    story: from.story,
    blockId: holder.paraId ?? to.blockId,
    offset,
    zeroWidthBefore:
      zeroWidthLeavesAt(target.content, offset).length - following,
  };
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

/**
 * Plan a tracked replacement of a same-list range: the range's tracked
 * deletion followed by the replacement fragment as tracked insertions.
 * All physical revision ids come from one pool for the complete batch.
 */
export const planTrackedReplace = (
  document: Document,
  options: PlanTrackedReplaceOptions,
): Result<DocumentOp[], DocumentOpRefusal> => {
  const { replacement, from, revision, newIds } = options;
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
  const deleted = appendTrackedDeletion({ document, options, plan });
  if (deleted.isErr()) return Result.err(deleted.error);
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
    const source = storyParagraphs(storyBody(plan.document(), from.story)).find(
      ({ paragraph: sourceParagraph }) =>
        idKey(sourceParagraph.paraId ?? "") === idKey(sourceBlockId),
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
  const toParagraph = storyParagraphs(storyBody(document, from.story)).find(
    ({ paragraph }) =>
      idKey(paragraph.paraId ?? "") === idKey(options.to.blockId),
  )?.paragraph;
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
    const first = storyParagraphs(storyBody(plan.document(), from.story)).find(
      ({ paragraph }) =>
        idKey(paragraph.paraId ?? "") === idKey(leadingFragment.paraId ?? ""),
    )?.paragraph;
    const payload = first?.content.at(-1);
    if (
      first !== undefined &&
      payload?.type === "insertion" &&
      payload.info.author === revision.author &&
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
      at,
      slice: replacement.tail,
      ...(options.seamPolicy === undefined
        ? {}
        : { seamPolicy: options.seamPolicy }),
      revision,
    });
    if (inserted.isErr()) return Result.err(inserted.error);
  }
  return Result.ok(plan.ops);
};

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
    const piece = order
      .slice(index + 1)
      .find(
        ({ node }) =>
          sameOwnFields(node, first) &&
          identitySlots(node).length === source.length &&
          identitySlots(node).every((slot) => !present.has(slotKey(slot))),
      )?.node;
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
