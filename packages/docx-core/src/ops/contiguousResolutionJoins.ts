import type { Result } from "better-result";

import type { Document, Paragraph } from "../model/document";
import type { DocumentEditEffects } from "./edits";
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

const isEmptyJoinParagraph = (
  paragraph: Paragraph,
): paragraph is IdentifiedParagraph =>
  paragraph.paraId !== undefined &&
  isParaId(paragraph.paraId) &&
  paragraph.content.length === 0 &&
  Object.keys(paragraph).every((key) => plainFields.has(key));

type PlanPairOptions = { join: JoinPlan; first: Paragraph; second: Paragraph };

type ContiguousJoinOptions = {
  document: Document;
  story: JoinPlan["story"];
  paragraphCount: number;
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
      paragraphs: Paragraph[];
      effects: DocumentEditEffects[];
      receiverEffects: ResolutionReceiverEffect[];
    };

/**
 * Each rejected inserted mark joins its paragraph into the already settled
 * suffix. For a chain of empty, section-free paragraphs, only the left paragraph
 * and the current survivor can affect the ordinary adjacent-pair plan.
 */
export const planContiguousResolutionJoins = ({
  document,
  story,
  paragraphCount,
  paragraphAt,
  joins,
  planPair,
}: ContiguousJoinOptions): ContiguousJoinResult => {
  let survivor = paragraphAt(paragraphCount - 1);
  if (
    story !== "main" ||
    joins.length === 0 ||
    joins.length + 1 !== paragraphCount ||
    survivor === undefined ||
    survivor.pPrMark !== undefined ||
    !isEmptyJoinParagraph(survivor)
  )
    return { type: "fallback", reason: "joinShape" };

  const effects: DocumentEditEffects[] = [];
  const receiverEffects: ResolutionReceiverEffect[] = [];
  for (let i = joins.length - 1; i >= 0; i--) {
    const first = paragraphAt(i);
    const join = joins[i];
    if (
      first === undefined ||
      !isEmptyJoinParagraph(first) ||
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
      !isEmptyJoinParagraph(next) ||
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
  return { type: "planned", paragraphs: [survivor], effects, receiverEffects };
};
