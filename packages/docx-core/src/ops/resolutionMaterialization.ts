import type { Document, DocumentBody, Paragraph } from "../model/document";
import { type ParagraphLocation, replaceParagraphs } from "./blocks";
import type { planContiguousResolutionJoins } from "./contiguousResolutionJoins";
import {
  hasMultipleParagraphPropertyChanges,
  meetsContract,
  repeatedParagraphId,
  validateOpsDocument,
} from "./contract";
import {
  combineEditEffects,
  type DocumentEdit,
  type DocumentEditEffects,
} from "./edits";
import { structurallyEqual } from "./equality";
import { idKey, isParaId } from "./ids";
import type { InlineNode } from "./leaves";
import { prepareResolutionParagraphOp } from "./paragraphPrimitive";
import { transferParagraphPropertySnapshot } from "./paragraphPropertySnapshot";
import type { ResolutionReceiverEffect } from "./resolutionMarkPlan";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type OpStory,
  type ReplaceInlineOp,
  type SetParagraphReviewOp,
} from "./types";
import { captureDocumentOp } from "./wire";

// Remove with the size-independent native paragraph splice. The older
// splice spreads its replacement into a function call.
const MAX_INLINE_RESOLUTION_SPAN = 16_384;

export type ResolutionInlineOp = ReplaceInlineOp | SetParagraphReviewOp;

type JoinResult = ReturnType<typeof planContiguousResolutionJoins>;
type JoinFallback = Extract<JoinResult, { type: "fallback" }>["reason"];

type JoinPreparation =
  | { type: "noJoins" }
  | {
      type: "plan";
      prepare: (view: {
        paragraphCount: number;
        indexOf: (id: string) => number | undefined;
        paragraphAt: (index: number) => Paragraph | undefined;
      }) => JoinResult;
    };

type MaterializationOptions = {
  document: Document;
  story: OpStory;
  locations: readonly ParagraphLocation[];
  inline: readonly ResolutionInlineOp[];
  join: JoinPreparation;
  noteReferences: readonly unknown[];
  receiverEffects: readonly ResolutionReceiverEffect[];
};

type FallbackReason =
  | JoinFallback
  | "story"
  | "packageDependencies"
  | "sectionView"
  | "paragraphList"
  | "referenceEffects"
  | "inputContract"
  | "primitiveShape"
  | "primitiveValidation"
  | "identityIntroduction"
  | "wideSpan"
  | "materialization";

type MaterializationResult =
  | { type: "fallback"; reason: FallbackReason }
  | {
      type: "materialized";
      edit: DocumentEdit;
      receiverEffects: readonly ResolutionReceiverEffect[];
    };

const fallback = (reason: FallbackReason): MaterializationResult => ({
  type: "fallback",
  reason,
});

const isTextRun = (node: InlineNode): boolean =>
  node.type === "run" && node.content.every(({ type }) => type === "text");

const isLocalText = (node: InlineNode): boolean =>
  isTextRun(node) ||
  ((node.type === "insertion" || node.type === "deletion") &&
    node.content.every(isTextRun));

// Called after ruling out physical section breaks. A projected body can
// already be canonical before a native splice has recorded that fact.
const hasCanonicalTrailingSection = (body: DocumentBody): boolean => {
  if (body.sections === undefined) return true;
  if (body.sections?.length !== 1 || body.finalSectionProperties === undefined)
    return false;
  const section = body.sections.at(0);
  return (
    section !== undefined &&
    section.properties === body.finalSectionProperties &&
    Object.keys(section).length === 2 &&
    Object.hasOwn(section, "properties") &&
    Object.hasOwn(section, "content") &&
    section.content.length === body.content.length &&
    body.content.every((block, index) => block === section.content[index])
  );
};

/**
 * The ordinary primitive validators prepare each paragraph in journal order.
 * They can read the original package here because each edit only retires IDs.
 * Section, note, range and other container interactions stay in the ordinary
 * sequential planner. A failed proof lets that planner determine the refusal.
 */
export const tryMaterializeResolution = ({
  document,
  story,
  locations,
  inline,
  join,
  noteReferences,
  receiverEffects,
}: MaterializationOptions): MaterializationResult => {
  if (story !== OP_STORIES.MAIN) return fallback("story");
  const body = document.package.document;
  if (
    (document.package.footnotes?.length ?? 0) > 0 ||
    (document.package.endnotes?.length ?? 0) > 0 ||
    (body.comments?.length ?? 0) > 0
  )
    return fallback("packageDependencies");
  if (noteReferences.length > 0 || receiverEffects.length > 0)
    return fallback("referenceEffects");
  const originals = body.content;
  if (
    !originals.every(
      (block): block is Paragraph => block.type === "paragraph",
    ) ||
    locations.length !== originals.length
  )
    return fallback("paragraphList");
  if (
    originals.some(
      ({ sectionProperties }) => sectionProperties !== undefined,
    ) ||
    !hasCanonicalTrailingSection(body)
  )
    return fallback("sectionView");
  // Normal operations have already validated the input. Callers inspecting a
  // plan directly still meet the same contract before using its effects.
  if (validateOpsDocument(document).isErr()) return fallback("inputContract");

  const byId = new Map<string, ParagraphLocation>();
  for (const at of locations) {
    const { paraId } = at.paragraph;
    if (
      at.list.length > 0 ||
      originals.at(at.index) !== at.paragraph ||
      paraId === undefined ||
      !isParaId(paraId)
    )
      return fallback("paragraphList");
    byId.set(idKey(paraId), at);
  }
  const replacements = new Map<number, Paragraph>();
  const effects: DocumentEditEffects[] = [];
  let first = originals.length;
  let end = 0;
  for (const op of inline) {
    const at = byId.get(idKey(op.blockId));
    const paragraph =
      at === undefined
        ? undefined
        : (replacements.get(at.index) ?? originals.at(at.index));
    if (
      op.story !== story ||
      at === undefined ||
      paragraph === undefined ||
      !paragraph.content.every(isLocalText)
    )
      return fallback("primitiveShape");
    if (op.type === DOCUMENT_OP_TYPES.REPLACE_INLINE) {
      if (!op.content.every(isTextRun)) return fallback("primitiveShape");
    } else if (
      op.review.propertyChanges !== undefined ||
      op.review.pPrMark !== undefined
    ) {
      return fallback("primitiveShape");
    }
    const prepared = prepareResolutionParagraphOp({
      document,
      op,
      at: { ...at, paragraph },
    });
    if (prepared.isErr()) return fallback("primitiveValidation");
    if (prepared.value.type === "unchanged") continue;
    const next = prepared.value;
    if (
      next.incomingParagraphIds.length > 0 ||
      next.incomingIdentityKeys.length > 0
    )
      return fallback("identityIntroduction");
    if (!structurallyEqual(next.paragraph.formatting, paragraph.formatting))
      return fallback("primitiveShape");
    transferParagraphPropertySnapshot({
      source: paragraph,
      target: next.paragraph,
    });
    replacements.set(at.index, next.paragraph);
    first = Math.min(first, at.index);
    end = Math.max(end, at.index + 1);
    effects.push({
      inverse: [captureDocumentOp(next.inverse)],
      touched: next.touched,
    });
  }

  if (effects.length === 0 && join.type === "noJoins") {
    return {
      type: "materialized",
      edit: combineEditEffects(document, []),
      receiverEffects: [],
    };
  }
  const combinedInline = combineEditEffects(document, effects);
  // The join planner reads the staged overlay directly. Only the no-join
  // branch needs to copy the span that the immutable splice will replace.
  const joined =
    join.type === "plan"
      ? join.prepare({
          paragraphCount: originals.length,
          indexOf: (id) => byId.get(idKey(id))?.index,
          paragraphAt: (index) =>
            replacements.get(index) ?? originals.at(index),
        })
      : undefined;
  if (joined?.type === "fallback") return joined;
  const start = joined === undefined ? first : Math.min(first, joined.start);
  const stop = joined === undefined ? end : Math.max(end, joined.end);
  const removed = joined === undefined ? 0 : joined.end - joined.start - 1;
  if (stop - start - removed > MAX_INLINE_RESOLUTION_SPAN)
    return fallback("wideSpan");
  let replacement: Paragraph[];
  if (joined === undefined) {
    replacement = originals.slice(start, stop);
    for (const [index, paragraph] of replacements)
      replacement[index - start] = paragraph;
  } else {
    replacement = [];
    // Keep inline work outside the join in the same physical splice. The
    // settled chain contributes one paragraph regardless of its length.
    for (let index = start; index < stop; index++) {
      if (index === joined.start) {
        replacement.push(joined.survivor);
        index = joined.end - 1;
      } else {
        const current = replacements.get(index) ?? originals.at(index);
        if (current === undefined) return fallback("paragraphList");
        replacement.push(current);
      }
    }
  }
  const final = replaceParagraphs({
    document,
    story,
    at: { list: [], index: start },
    count: stop - start,
    replacement,
  });
  if (
    final.isErr() ||
    hasMultipleParagraphPropertyChanges(final.value) ||
    repeatedParagraphId(final.value) !== undefined
  )
    return fallback("materialization");
  meetsContract(final.value);
  return {
    type: "materialized",
    edit: combineEditEffects(final.value, [
      combinedInline,
      ...(joined?.effects ?? []),
    ]),
    receiverEffects: joined?.receiverEffects ?? [],
  };
};
