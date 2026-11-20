import { Result } from "better-result";

import type { Document, Paragraph } from "../model/document";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import { endsItsContainer, type ParagraphLocation, storyBody } from "./blocks";
import type { DocumentEditEffects } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import {
  countIds,
  countKeys,
  identityKeysIn,
  idKey,
  isParaId,
  type KeyMembership,
  packageIdentityKeyIndex,
  packageParagraphIdIndex,
  paragraphIdsIn,
} from "./ids";
import {
  childNodes,
  type InlineNode,
  isEmptyRecord,
  leafSpans,
  textsIn,
} from "./leaves";
import { cloneModel } from "./modelClone";
import { automaticNoteMarkRuns } from "./noteMarks";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
import { reviewFieldsOf, withReviewFields } from "./review";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type OpStory,
  type ParagraphReviewFields,
  type ReplaceBlocksOp,
  type ReplaceInlineOp,
  SECTION_BOUNDARY_POLICIES,
  type SetParagraphReviewOp,
  type TouchedBlocks,
} from "./types";

const refusal = (
  op: DocumentOp,
  reason: DocumentOpRefusalReason,
  message: string,
): DocumentOpRefusal =>
  new DocumentOpRefusal({ message, reason, opType: op.type });

export const holdsEmptyRecord = (nodes: readonly InlineNode[]): boolean =>
  nodes.some(
    (node) => isEmptyRecord(node) || holdsEmptyRecord(childNodes(node) ?? []),
  );

export const withContent = (
  paragraph: Paragraph,
  content: readonly Paragraph["content"][number][],
): Paragraph =>
  content === paragraph.content
    ? paragraph
    : { ...paragraph, content: [...content] };

/** The replacement that gives a paragraph back the content it had. */
export const contentRestoring = (
  story: OpStory,
  result: Paragraph,
  original: Paragraph,
): ReplaceInlineOp => ({
  type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
  story,
  blockId: result.paraId ?? original.paraId ?? "",
  expected: result.content,
  content: original.content,
});

/** The review fields' setting that gives a paragraph back the ones it had. */
export const reviewRestoring = (
  story: OpStory,
  result: Paragraph,
  original: Paragraph,
): SetParagraphReviewOp => ({
  type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
  story,
  blockId: result.paraId ?? original.paraId ?? "",
  expected: reviewFieldsOf(result),
  review: reviewFieldsOf(original),
});

export const touchedBetween = (
  before: readonly Paragraph[],
  after: readonly Paragraph[],
): TouchedBlocks => {
  const beforeIds = before.flatMap(({ paraId }) =>
    paraId === undefined ? [] : [paraId],
  );
  const afterIds = after.flatMap(({ paraId }) =>
    paraId === undefined ? [] : [paraId],
  );
  const afterKeys = new Set(afterIds.map(idKey));
  const beforeKeys = new Set(beforeIds.map(idKey));
  return {
    modified: beforeIds.filter((id) => afterKeys.has(idKey(id))),
    inserted: afterIds.filter((id) => !beforeKeys.has(idKey(id))),
    removed: beforeIds.filter((id) => !afterKeys.has(idKey(id))),
  };
};

/** Revision and control ids the package uses outside `before`. */
export const identitiesOutside = (
  document: Document,
  before: readonly Paragraph[],
): KeyMembership => {
  const index = packageIdentityKeyIndex(document.package);
  const inside = countKeys(identityKeysIn(before));
  return { has: (key) => index.count(key) > (inside.get(key) ?? 0) };
};

type ParagraphRecordCollisionOptions = {
  document: Document;
  paragraph: Paragraph;
  incoming: readonly string[];
  kept: () => readonly string[];
};

const paragraphRecordsCollide = ({
  document,
  paragraph,
  incoming,
  kept,
}: ParagraphRecordCollisionOptions): boolean => {
  if (incoming.length === 0) return false;
  const outside = identitiesOutside(document, [paragraph]);
  const own = new Set(kept());
  return (
    incoming.some((key) => outside.has(key) || own.has(key)) ||
    new Set(incoming).size !== incoming.length
  );
};

type ParagraphIdCollisionOptions = {
  document: Document;
  incoming: readonly string[];
  removed?: unknown;
};

export const paragraphIdsCollide = ({
  document,
  incoming,
  removed,
}: ParagraphIdCollisionOptions): boolean => {
  if (incoming.length === 0) return false;
  const index = packageParagraphIdIndex(document.package);
  const removedCounts = countIds(paragraphIdsIn(removed));
  const seen = new Set<string>();
  for (const id of incoming) {
    const key = idKey(id);
    if (seen.has(key) || index.count(key) > (removedCounts.get(key) ?? 0))
      return true;
    seen.add(key);
  }
  return false;
};

/**
 * Whether a replacement leaves every section break where it was: the replaced
 * paragraphs lie in one section (only the last may end it) and the
 * replacement ends it the same way. Anything else would change the sections
 * the body is divided into, which is a section operation.
 */
export const sameSectionBreaks = (
  before: readonly Paragraph[],
  after: readonly Paragraph[],
): boolean => {
  const breaksInside = (paragraphs: readonly Paragraph[]) =>
    paragraphs
      .slice(0, -1)
      .some(({ sectionProperties }) => sectionProperties !== undefined);
  return (
    !breaksInside(before) &&
    !breaksInside(after) &&
    structurallyEqual(
      before.at(-1)?.sectionProperties,
      after.at(-1)?.sectionProperties,
    )
  );
};

type ValidateBlockReplacementOptions = {
  document: Document;
  op: ReplaceBlocksOp;
  before: readonly Paragraph[];
};

export const validateBlockReplacement = ({
  document,
  op,
  before,
}: ValidateBlockReplacementOptions): Result<void, DocumentOpRefusal> => {
  if (op.blocks.some(({ paraId }) => paraId === undefined)) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        "A replacement paragraph has no id.",
      ),
    );
  }
  if (op.blocks.some(({ content }) => holdsEmptyRecord(content))) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT,
        "A replacement paragraph holds an empty run, text node, or revision wrapper.",
      ),
    );
  }
  // An id only one side has is created by the replacement or by its inverse.
  const kept = new Set(before.map(({ paraId }) => idKey(paraId ?? "")));
  const placed = new Set(op.blocks.map(({ paraId }) => idKey(paraId ?? "")));
  const created = [
    ...op.blocks.filter(({ paraId }) => !kept.has(idKey(paraId ?? ""))),
    ...before.filter(({ paraId }) => !placed.has(idKey(paraId ?? ""))),
  ];
  if (created.some(({ paraId }) => paraId === undefined || !isParaId(paraId))) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        "A paragraph the replacement adds or removes has no usable id.",
      ),
    );
  }
  if (
    op.sectionBoundaries !== SECTION_BOUNDARY_POLICIES.REPLACE &&
    !sameSectionBreaks(before, op.blocks)
  ) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
        "A replacement keeps every section break where it is.",
      ),
    );
  }
  if (
    paragraphIdsCollide({
      document,
      incoming: paragraphIdsIn(op.blocks),
      removed: before,
    })
  ) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
        "A replacement paragraph id is already used in the package.",
      ),
    );
  }
  const incomingRecords = identityKeysIn(op.blocks);
  if (incomingRecords.length > 0) {
    const outside = identitiesOutside(document, before);
    if (
      incomingRecords.some((key) => outside.has(key)) ||
      new Set(incomingRecords).size !== incomingRecords.length
    ) {
      return Result.err(
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
          "A replacement carries a revision or content-control id already used in the package.",
        ),
      );
    }
  }
  return Result.ok(undefined);
};

/** The primitive inverse and touched IDs of a nonempty block replacement. */
export const nonemptyBlockReplacementEffects = (
  op: ReplaceBlocksOp,
  before: readonly Paragraph[],
): DocumentEditEffects => ({
  inverse: [
    {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: op.story,
      expected: op.blocks,
      blocks: before,
      ...(op.sectionBoundaries === SECTION_BOUNDARY_POLICIES.REPLACE
        ? { sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE }
        : {}),
    },
  ],
  touched: touchedBetween(before, op.blocks),
});

export type ResolutionParagraphOp = SetParagraphReviewOp | ReplaceInlineOp;
type ResolutionParagraphPlan =
  | { type: "unchanged" }
  | {
      type: "replacement";
      paragraph: Paragraph;
      inverse: DocumentOp;
      touched: TouchedBlocks;
      incomingParagraphIds: readonly string[];
      incomingIdentityKeys: readonly string[];
    };
type Prepared = Result<ResolutionParagraphPlan, DocumentOpRefusal>;
type ParagraphPrimitiveOptions = {
  document: Document;
  op: ResolutionParagraphOp;
  at: ParagraphLocation;
};

const refuse = (
  op: ResolutionParagraphOp,
  reason: DocumentOpRefusalReason,
  message: string,
): Prepared => Result.err(refusal(op, reason, message));

/** Revision and control ids a paragraph's own review fields carry, as slot keys. */
const reviewKeysOf = (review: ParagraphReviewFields): string[] =>
  identityKeysIn({
    propertyChanges: review.propertyChanges,
    pPrMark: review.pPrMark,
  });

const prepareParagraphReview = ({
  document,
  op,
  at,
}: ParagraphPrimitiveOptions & { op: SetParagraphReviewOp }): Prepared => {
  const { paragraph } = at;
  if (
    !equalForStaleness(reviewFieldsOf(paragraph), reviewFieldsOf(op.expected))
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `${op.blockId} states other review fields than expected.`,
    );
  }
  const review = cloneModel(op.review);
  const next = withReviewFields(paragraph, review);
  if (structurallyEqual(next, paragraph)) {
    return Result.ok({ type: "unchanged" });
  }
  type TerminalMarkTransitionOptions = {
    from: Paragraph["pPrMark"];
    to: Paragraph["pPrMark"];
  };
  const refusesMarkTransition = ({ from, to }: TerminalMarkTransitionOptions) =>
    to !== undefined &&
    !structurallyEqual(to, from) &&
    endsItsContainer(storyBody(document, op.story), at) &&
    !(
      at.list.some((step) => step.kind === "tableCell") &&
      (to.kind === "ins" || to.kind === "del")
    );
  if (refusesMarkTransition({ from: paragraph.pPrMark, to: review.pPrMark })) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
      `${op.blockId} ends its container: there is no next paragraph its mark could join.`,
    );
  }
  const incomingIdentityKeys = reviewKeysOf(review);
  if (
    paragraphRecordsCollide({
      document,
      paragraph,
      incoming: incomingIdentityKeys,
      kept: () => identityKeysIn(paragraph.content),
    })
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "The review fields carry a revision id already used in the package.",
    );
  }
  return Result.ok({
    type: "replacement",
    paragraph: next,
    // Imported terminal marks may be resolved even when a new mark could
    // not be set there. Restore that exact source snapshot, with a whole
    // paragraph stale check, rather than replaying a prohibited new set.
    inverse: refusesMarkTransition({
      from: next.pPrMark,
      to: paragraph.pPrMark,
    })
      ? {
          type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
          story: op.story,
          expected: [next],
          blocks: [paragraph],
        }
      : reviewRestoring(op.story, next, paragraph),
    touched: touchedBetween([paragraph], [next]),
    incomingParagraphIds: [],
    incomingIdentityKeys,
  });
};

const prepareInlineReplacement = ({
  document,
  op,
  at,
}: ParagraphPrimitiveOptions & { op: ReplaceInlineOp }): Prepared => {
  const { paragraph } = at;
  if (!equalForStaleness(paragraph.content, op.expected)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `${op.blockId} holds other content than expected.`,
    );
  }
  let content = cloneModel(op.content);
  const noteKind =
    op.story !== "main" &&
    (op.story.kind === "footnote" || op.story.kind === "endnote")
      ? op.story.kind
      : undefined;
  if (
    noteKind !== undefined &&
    !leafSpans(content).some(
      ({ node }) => node.type === "noteMarker" && node.kind === noteKind,
    )
  ) {
    const markers = automaticNoteMarkRuns(paragraph.content).filter((run) =>
      run.content.some(
        (node) => node.type === "noteMarker" && node.kind === noteKind,
      ),
    );
    content = [...markers, ...content];
  }
  if (holdsEmptyRecord(content)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT,
      "The content holds an empty run, text node, or revision wrapper.",
    );
  }
  if (textsIn(content).some(hasIllegalXmlCharacters)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      "The content holds text that cannot be written.",
    );
  }
  if (structurallyEqual(content, paragraph.content)) {
    return Result.ok({ type: "unchanged" });
  }
  const incomingParagraphIds = paragraphIdsIn(content);
  if (
    paragraphIdsCollide({
      document,
      incoming: incomingParagraphIds,
      removed: paragraph.content,
    })
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "The content holds a paragraph id the package already uses.",
    );
  }
  const incomingIdentityKeys = identityKeysIn(content);
  if (
    paragraphRecordsCollide({
      document,
      paragraph,
      incoming: incomingIdentityKeys,
      kept: () => reviewKeysOf(reviewFieldsOf(paragraph)),
    })
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "The content carries a revision or content-control id already used in the package.",
    );
  }
  const next = withContent(paragraph, content);
  return Result.ok({
    type: "replacement",
    paragraph: next,
    inverse: contentRestoring(op.story, next, paragraph),
    touched: touchedBetween([paragraph], [next]),
    incomingParagraphIds,
    incomingIdentityKeys,
  });
};

/** Plan a native paragraph primitive before materializing its block list. */
export const prepareResolutionParagraphOp = ({
  document,
  op,
  at,
}: ParagraphPrimitiveOptions): Prepared => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
      return prepareParagraphReview({ document, op, at });
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
      return prepareInlineReplacement({ document, op, at });
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};
