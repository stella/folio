/** Tracked replacement fragments share deletion's validation and physical-id pool. */
import { Result, panic } from "better-result";

import type { Document, Paragraph } from "../model/document";
import { storyBody, storyParagraphs } from "./blocks";
import { identityKeysIn, idKey } from "./ids";
import { defaultInsertionGap } from "./leaves";
import { gapAfterInserted } from "./inline";
import { appendTrackedDeletion, createTrackedPlan, type PlanTrackedDeletionOptions } from "./plan";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  SPLIT_HALVES,
  type DocumentOp,
  type InlineSlice,
  type TextPosition,
} from "./types";

type RangeStartAfterDeletionOptions = {
  before: Document;
  after: Document;
  from: TextPosition;
  to: TextPosition;
};

/** Own inserted marks can be cancelled, moving the prefix to its next live paragraph. */
export const rangeStartAfterDeletion = ({
  before,
  after,
  from,
  to,
}: RangeStartAfterDeletionOptions): TextPosition => {
  const original = storyParagraphs(storyBody(before, from.story));
  const first = original.findIndex(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(from.blockId),
  );
  const last = original.findIndex(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(to.blockId),
  );
  const live = new Set(
    storyParagraphs(storyBody(after, from.story)).map(({ paragraph }) =>
      idKey(paragraph.paraId ?? ""),
    ),
  );
  const retained = original
    .slice(first, last + 1)
    .find(({ paragraph }) => live.has(idKey(paragraph.paraId ?? "")));
  if (retained === undefined)
    panic("A successful range deletion must retain a selected paragraph.");
  const source = original.at(first)?.paragraph;
  if (source === undefined) panic("A successful deletion must have its input paragraph.");
  return {
    ...from,
    blockId: retained.paragraph.paraId ?? from.blockId,
    zeroWidthBefore:
      from.zeroWidthBefore ?? defaultInsertionGap(source.content, from.offset).zeroWidthBefore,
  };
};

/**
 * A range replacement fragment: newly identified paragraphs before its last
 * paragraph, and the inline tail inserted into the retained original paragraph.
 * Each new paragraph's own fields apply to it, including when it takes the
 * source prefix. The last fragment keeps the original survivor's identity.
 * Closed inline slices are required; review and content-control ids in the
 * fragment, and section-bearing paragraphs, are unsupported.
 */
export type PlanTrackedReplaceOptions = PlanTrackedDeletionOptions & {
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
    identityKeysIn(replacement).length > 0 ||
    replacement.paragraphs.some(({ sectionProperties }) => sectionProperties !== undefined)
  ) {
    return refuse(
      "Tracked replacement cannot insert identified review, content-control or section records.",
    );
  }
  const plan = createTrackedPlan({ document, revision, newIds });
  const deleted = appendTrackedDeletion({ document, options, plan });
  if (deleted.isErr()) return Result.err(deleted.error);
  // Tracking leaves selected content in place. Every fragment is inserted at
  // its leading boundary; splitting moves the old mark to the original half.
  let at = rangeStartAfterDeletion({
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
    if (source === undefined) return refuse("The replacement's source paragraph does not exist.");
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
    at = { story: from.story, blockId: at.blockId, offset: 0, zeroWidthBefore: 0 };
  }
  if (replacement.tail.content.length > 0) {
    const inserted = plan.append({
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at,
      slice: replacement.tail,
      ...(options.seamPolicy === undefined ? {} : { seamPolicy: options.seamPolicy }),
      revision,
    });
    if (inserted.isErr()) return Result.err(inserted.error);
  }
  return Result.ok(plan.ops);
};
