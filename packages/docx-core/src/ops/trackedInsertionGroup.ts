import type { Document, Paragraph } from "../model/document";
import { prepareTrackedParagraphInsertion } from "./blockInsertion";
import { replaceParagraphs } from "./blocks";
import {
  hasMultipleParagraphPropertyChanges,
  meetsContract,
  repeatedParagraphId,
  validateOpsDocument,
} from "./contract";
import type { DocumentEditEffects } from "./edits";
import {
  type IdentityKeyIndex,
  identityKeysIn,
  idKey,
  isParaId,
  packageIdentityKeyIndex,
  packageParagraphIdIndex,
  paragraphIdsIn,
} from "./ids";
import type { InlineNode } from "./leaves";
import { transferParagraphPropertySnapshot } from "./paragraphPropertySnapshot";
import type { DocumentOpRefusal } from "./refusal";
import { stampedRevisionIds } from "./review";
import { trackedStampRefusal } from "./trackedStamp";
import { hasExactTrailingSection } from "./trailingSectionMirror";
import { DOCUMENT_OP_TYPES, type InsertBlocksOp, OP_STORIES } from "./types";
import { captureDocumentOp, restoreDocumentOp } from "./wire";

// Remove when the shared paragraph splice stops using variadic arguments.
// Each replacement also includes the original anchor.
const MAX_GROUP_MEMBERS = 16_383;

type WorkingIndex = Pick<IdentityKeyIndex, "count" | "has"> & {
  replace: (before: readonly string[], after: readonly string[]) => void;
};

const withLocalCounts = (
  initial: Pick<IdentityKeyIndex, "count">,
): WorkingIndex => {
  const changes = new Map<string, number>();
  const count = (key: string) => initial.count(key) + (changes.get(key) ?? 0);
  return {
    count,
    has: (key) => count(key) > 0,
    replace: (before, after) => {
      for (const key of before) changes.set(key, (changes.get(key) ?? 0) - 1);
      for (const key of after) changes.set(key, (changes.get(key) ?? 0) + 1);
    },
  };
};

type MemberEffect = DocumentEditEffects & {
  op: InsertBlocksOp;
  revisions: readonly number[];
};

type FallbackReason =
  | "compiler"
  | "sourceShape"
  | "sectionView"
  | "memberShape"
  | "effectShape"
  | "wideSpan"
  | "materialization";

type TrackedInsertionGroupResult =
  | { type: "fallback"; reason: FallbackReason }
  | {
      type: "fallback";
      reason: "refusal";
      index: number;
      refusal: DocumentOpRefusal;
    }
  | {
      type: "materialized";
      document: Document;
      members: readonly MemberEffect[];
    };

const fallback = (reason: FallbackReason): TrackedInsertionGroupResult => ({
  type: "fallback",
  reason,
});

const isTextRun = (node: InlineNode): boolean =>
  node.type === "run" && node.content.every(({ type }) => type === "text");

const hasOnlyLocalText = (node: InlineNode): boolean =>
  isTextRun(node) ||
  ((node.type === "insertion" || node.type === "deletion") &&
    node.content.every(isTextRun));

type NextMember =
  | { type: "member"; op: InsertBlocksOp }
  | { type: "done" }
  | { type: "fallback" };

type TrackedInsertionGroupOptions = {
  document: Document;
  next: (previousRevisions: readonly number[] | undefined) => NextMember;
};

/**
 * Attempt a precompiled journal run at one physical main-story anchor. The
 * prepared members see all prior ids and the real current source/follower.
 * Publish a document and journal effects only when the whole group succeeds.
 * Keep the source document and all emitted operations immutable during the
 * synchronous compiler callback. The caller owns its journal copies.
 */
export const tryTrackedInsertionGroup = ({
  document,
  next,
}: TrackedInsertionGroupOptions): TrackedInsertionGroupResult => {
  const firstMember = next(undefined);
  if (firstMember.type === "fallback") return fallback("compiler");
  if (firstMember.type !== "member") return fallback("memberShape");
  const first = firstMember.op;
  if (first.at.type !== "before" && first.at.type !== "after")
    return fallback("memberShape");
  const anchorKey = idKey(first.at.blockId);
  const side = first.at.type;
  const body = document.package.document;
  if (
    (document.package.footnotes?.length ?? 0) > 0 ||
    (document.package.endnotes?.length ?? 0) > 0 ||
    (body.comments?.length ?? 0) > 0 ||
    !body.content.every(
      (block): block is Paragraph =>
        block.type === "paragraph" && block.sectionProperties === undefined,
    )
  )
    return fallback("sourceShape");
  if (body.sections !== undefined && !hasExactTrailingSection(body))
    return fallback("sectionView");
  if (validateOpsDocument(document).isErr()) return fallback("sourceShape");
  const anchorAt = body.content.findIndex(
    ({ paraId }) => paraId !== undefined && idKey(paraId) === anchorKey,
  );
  if (anchorAt < 0) return fallback("sourceShape");
  const source = body.content.at(anchorAt);
  if (
    source === undefined ||
    source.paraId === undefined ||
    !isParaId(source.paraId) ||
    !source.content.every(hasOnlyLocalText)
  )
    return fallback("sourceShape");

  const paragraphIds = withLocalCounts(
    packageParagraphIdIndex(document.package),
  );
  const recordIds = withLocalCounts(packageIdentityKeyIndex(document.package));
  let anchor = source;
  let following = body.content.at(anchorAt + 1);
  const inserted: Paragraph[] = [];
  const members: MemberEffect[] = [];
  let member: NextMember = firstMember;
  while (member.type === "member") {
    const input = member.op;
    const index = members.length;
    if (index >= MAX_GROUP_MEMBERS) return fallback("wideSpan");
    const restored = restoreDocumentOp(input);
    if (restored.isErr())
      return {
        type: "fallback",
        reason: "refusal",
        index,
        refusal: restored.error,
      };
    const op = restored.value;
    if (
      op.type !== DOCUMENT_OP_TYPES.INSERT_BLOCKS ||
      op.story !== OP_STORIES.MAIN ||
      op.revision === undefined ||
      op.at.type !== side ||
      idKey(op.at.blockId) !== anchorKey ||
      op.move === true ||
      op.blocks.length !== 1
    )
      return fallback("memberShape");
    const badStamp = trackedStampRefusal(op, op.revision, recordIds.has);
    if (badStamp !== undefined)
      return { type: "fallback", reason: "refusal", index, refusal: badStamp };
    const prepared = prepareTrackedParagraphInsertion({
      document,
      op,
      anchor,
      ...(following === undefined ? {} : { following }),
      validatedNeighbor: side === "after" ? (following ?? anchor) : anchor,
      paragraphIds,
      recordIds,
    });
    if (prepared.isErr())
      return {
        type: "fallback",
        reason: "refusal",
        index,
        refusal: prepared.error,
      };
    if (prepared.value.type !== "replacement") return fallback("memberShape");
    const { before, operation, effect } = prepared.value;
    const [left, right] = operation.blocks;
    const takesOverMark = side === "after" && following === undefined;
    const retained = takesOverMark ? left : right;
    const added = takesOverMark ? right : left;
    if (
      operation.story !== op.story ||
      operation.sectionBoundaries !== undefined ||
      operation.sectionView !== undefined ||
      operation.expected.length !== 1 ||
      operation.expected.at(0) !== before ||
      operation.blocks.length !== 2 ||
      retained === undefined ||
      added === undefined ||
      before !== (side === "after" ? (following ?? anchor) : anchor) ||
      (takesOverMark
        ? retained.paraId !== before.paraId ||
          retained.content !== before.content ||
          retained.formatting !== before.formatting ||
          retained.sectionProperties !== undefined
        : retained !== before) ||
      added.sectionProperties !== undefined ||
      added.paraId === undefined ||
      paragraphIds.count(idKey(added.paraId)) !== 0 ||
      (added.propertyChanges?.length ?? 0) > 1
    )
      return fallback("effectShape");

    const touchedIds = new Set(
      [...effect.touched.modified, ...effect.touched.inserted].map(idKey),
    );
    const revisions = stampedRevisionIds(
      operation.blocks.filter(
        ({ paraId }) => paraId !== undefined && touchedIds.has(idKey(paraId)),
      ),
      op.revision,
      recordIds,
    );
    if (takesOverMark) {
      transferParagraphPropertySnapshot({ source: anchor, target: retained });
      anchor = retained;
    }
    inserted.push(added);
    if (side === "after") following = added;
    // Counts on an identical retained paragraph cancel in both identity spaces.
    const removed = retained === before ? [] : [before];
    const incoming = retained === before ? [added] : operation.blocks;
    paragraphIds.replace(
      paragraphIdsIn(removed).map(idKey),
      paragraphIdsIn(incoming).map(idKey),
    );
    recordIds.replace(identityKeysIn(removed), identityKeysIn(incoming));
    members.push({
      op: input,
      inverse: effect.inverse.map(captureDocumentOp),
      touched: effect.touched,
      revisions,
    });
    member = next(revisions);
  }
  if (member.type === "fallback") return fallback("compiler");

  if (side === "after") inserted.reverse();
  const replacement =
    side === "before" ? [...inserted, anchor] : [anchor, ...inserted];
  const final = replaceParagraphs({
    document,
    story: OP_STORIES.MAIN,
    at: { list: [], index: anchorAt },
    count: 1,
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
    document: final.value,
    members,
  };
};
