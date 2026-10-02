import { Result } from "better-result";

import type { Document } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyBody, storyParagraphs } from "../blocks";
import { paragraphLength } from "../offsets";
import type { DocumentOpRefusal } from "../refusal";
import { DOCUMENT_OP_TYPES, SPLIT_HALVES, type DocumentOp } from "../types";
import { NO_SEQUENCED_EFFECT, type SequencedBatch } from "./envelope";

const TOUCHED_BLOCKS_EFFECT = "touchedBlocks";

type AppliedBatch = {
  document: Document;
  inverse: readonly DocumentOp[];
  effects: NonNullable<SequencedBatch["effects"]>;
};

/** Capture position facts before each operation, including within an atomic batch. */
export const applyBatch = (
  document: Document,
  ops: readonly DocumentOp[],
): Result<AppliedBatch, DocumentOpRefusal> => {
  let current = document;
  const inverses: (readonly DocumentOp[])[] = [];
  const effects: NonNullable<SequencedBatch["effects"]>[number][] = [];
  for (const op of ops) {
    let effect: NonNullable<SequencedBatch["effects"]>[number] = { type: NO_SEQUENCED_EFFECT };
    if (op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK) {
      const paragraph = storyParagraphs(storyBody(current, op.at.story)).find(
        ({ paragraph: candidate }) => candidate.paraId === op.at.blockId,
      )?.paragraph;
      if (paragraph !== undefined) {
        effect = {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          newHalf:
            op.newHalf ??
            (op.at.offset === paragraphLength(paragraph)
              ? SPLIT_HALVES.SECOND
              : SPLIT_HALVES.FIRST),
        };
      }
    }
    if (op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS) {
      const paragraph = storyParagraphs(storyBody(current, op.story)).find(
        ({ paragraph: candidate }) => candidate.paraId === op.blockId,
      )?.paragraph;
      if (paragraph !== undefined) {
        effect = { type: DOCUMENT_OP_TYPES.JOIN_BLOCKS, firstLength: paragraphLength(paragraph) };
      }
    }
    const applied = applyDocumentOps(current, [op]);
    if (applied.isErr()) {
      return Result.err(applied.error);
    }
    if (
      op.type === DOCUMENT_OP_TYPES.DELETE_BLOCKS ||
      op.type === DOCUMENT_OP_TYPES.RESOLVE_REVISION
    ) {
      const { modified, inserted, removed } = applied.value.touched;
      effect = {
        type: TOUCHED_BLOCKS_EFFECT,
        blockIds: [...new Set([...modified, ...inserted, ...removed])],
      };
    }
    current = applied.value.document;
    inverses.push(applied.value.inverse);
    effects.push(effect);
  }
  return Result.ok({ document: current, inverse: inverses.reverse().flat(), effects });
};
