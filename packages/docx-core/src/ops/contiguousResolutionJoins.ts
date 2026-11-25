import type { Result } from "better-result";

import type { Document, Paragraph } from "../model/document";
import type { DocumentEditEffects } from "./edits";
import { structurallyEqual } from "./equality";
import { isParaId } from "./ids";
import {
  nonemptyBlockReplacementEffects,
  validateBlockReplacement,
} from "./paragraphPrimitive";
import { transferParagraphPropertySnapshot } from "./paragraphPropertySnapshot";
import type { DocumentOpRefusal } from "./refusal";
import type {
  JoinPlan,
  ParagraphMarkPlan,
  ResolutionReceiverEffect,
} from "./resolutionMarkPlan";
import { DOCUMENT_OP_TYPES, REVISION_DECISIONS } from "./types";
import { captureDocumentOp } from "./wire";

const plainFields = new Set(["type", "paraId", "textId", "content", "pPrMark"]);

type IdentifiedParagraph = Paragraph & { paraId: string };

const isPlainTextContent = (paragraph: Paragraph): boolean => {
  if (paragraph.content.length === 0) return true;
  const run = paragraph.content.at(0);
  if (
    paragraph.content.length !== 1 ||
    run?.type !== "run" ||
    Object.keys(run).some((key) => key !== "type" && key !== "content") ||
    run.content.length !== 1
  )
    return false;
  const text = run.content.at(0);
  return (
    text?.type === "text" &&
    Object.keys(text).every((key) => key === "type" || key === "text")
  );
};

const isPlainJoinParagraph = (
  paragraph: Paragraph,
): paragraph is IdentifiedParagraph =>
  paragraph.paraId !== undefined &&
  isParaId(paragraph.paraId) &&
  isPlainTextContent(paragraph) &&
  Object.keys(paragraph).every((key) => plainFields.has(key));

type PlanPairOptions = { join: JoinPlan; first: Paragraph; second: Paragraph };

type ContiguousJoinOptions = {
  document: Document;
  story: JoinPlan["story"];
  paragraphCount: number;
  indexOf: (id: string) => number | undefined;
  paragraphAt: (index: number) => Paragraph | undefined;
  joins: readonly JoinPlan[];
  planPair: (
    pair: PlanPairOptions,
  ) => Result<ParagraphMarkPlan, DocumentOpRefusal>;
};

type ContiguousJoinResult =
  | { type: "fallback"; reason: "joinShape" | "joinPlan" | "joinValidation" }
  | {
      type: "planned";
      start: number;
      end: number;
      survivor: Paragraph;
      effects: DocumentEditEffects[];
      receiverEffects: ResolutionReceiverEffect[];
    };

/**
 * Each rejected inserted mark joins its paragraph into the already settled
 * suffix of its contiguous run. In a section-free pair with at most one plain
 * text payload, only those two paragraphs can affect the ordinary join plan.
 */
export const planContiguousResolutionJoins = ({
  document,
  story,
  paragraphCount,
  indexOf,
  paragraphAt,
  joins,
  planPair,
}: ContiguousJoinOptions): ContiguousJoinResult => {
  const firstJoin = joins.at(0);
  const start = firstJoin === undefined ? undefined : indexOf(firstJoin.paraId);
  if (start === undefined) return { type: "fallback", reason: "joinShape" };
  const end = start + joins.length + 1;
  let survivor = paragraphAt(end - 1);
  if (
    story !== "main" ||
    end > paragraphCount ||
    survivor === undefined ||
    survivor.pPrMark !== undefined ||
    !isPlainJoinParagraph(survivor)
  )
    return { type: "fallback", reason: "joinShape" };

  const effects: DocumentEditEffects[] = [];
  const receiverEffects: ResolutionReceiverEffect[] = [];
  for (let i = joins.length - 1; i >= 0; i--) {
    const first = paragraphAt(start + i);
    const join = joins[i];
    if (
      first === undefined ||
      !isPlainJoinParagraph(first) ||
      (first.content.length > 0 && survivor.content.length > 0) ||
      join === undefined ||
      join.story !== story ||
      join.op.decision !== REVISION_DECISIONS.REJECT ||
      !join.added ||
      join.depth !== undefined ||
      join.paraId !== first.paraId ||
      first.pPrMark?.kind !== "ins" ||
      first.pPrMark.info.id !== join.revisionId ||
      first.pPrMark.deletion !== undefined ||
      first.pPrMark.resolutionJoin !== undefined
    )
      return { type: "fallback", reason: "joinShape" };
    const second: Paragraph = survivor;
    const planned = planPair({ join, first, second });
    if (planned.isErr()) return { type: "fallback", reason: "joinPlan" };
    const plan = planned.value;
    const op = plan.ops.at(0);
    const receiver = plan.receiverEffects.at(0);
    if (
      plan.endsControl !== undefined ||
      plan.ops.length !== 1 ||
      op?.type !== DOCUMENT_OP_TYPES.REPLACE_BLOCKS ||
      op.story !== story ||
      op.sectionView !== undefined ||
      op.sectionBoundaries !== undefined ||
      op.expected.length !== 2 ||
      op.expected[0] !== first ||
      op.expected[1] !== second ||
      op.blocks.length !== 1 ||
      plan.receiverEffects.length !== 1 ||
      receiver?.type !== "paragraph-join" ||
      receiver.receiver !== second
    )
      return { type: "fallback", reason: "joinPlan" };

    const next = op.blocks[0];
    if (
      next === undefined ||
      next.pPrMark !== undefined ||
      !isPlainJoinParagraph(next) ||
      !structurallyEqual(
        next.content,
        first.content.length === 0 ? second.content : first.content,
      ) ||
      (next.paraId !== first.paraId && next.paraId !== second.paraId)
    )
      return { type: "fallback", reason: "joinPlan" };
    // Each step retains one source id and introduces no record identity.
    // The untouched package therefore gives exactly the same collision verdict.
    const valid = validateBlockReplacement({
      document,
      op,
      before: op.expected,
    });
    if (valid.isErr()) return { type: "fallback", reason: "joinValidation" };
    const survivorId = next.paraId;
    const source = next.paraId === first.paraId ? first : second;
    if (next !== source)
      transferParagraphPropertySnapshot({ source, target: next });
    const effect = nonemptyBlockReplacementEffects(op, op.expected);
    effects.push({
      inverse: effect.inverse.map(captureDocumentOp),
      touched: {
        ...effect.touched,
        absorbed: effect.touched.removed.map((absorbed) => ({
          absorbed,
          survivor: survivorId,
        })),
      },
    });
    receiverEffects.push(receiver);
    survivor = next;
  }
  return { type: "planned", start, end, survivor, effects, receiverEffects };
};
