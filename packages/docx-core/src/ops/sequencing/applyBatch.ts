import { Result } from "better-result";

import type { Document } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyBody, storyParagraphs } from "../blocks";
import { IDENTITY_SPACES, identityKeysIn, idKey, slotKey } from "../ids";
import { runGaps } from "../leaves";
import { paragraphLength } from "../offsets";
import type { DocumentOpRefusal } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertContentOp,
  type SetRunPropsOp,
  SPLIT_HALVES,
} from "../types";
import {
  NO_SEQUENCED_EFFECT,
  type SequencedBatch,
  sequencedEffectType,
} from "./envelope";
import { blockScopedEffect, TOUCHED_BLOCKS_EFFECT } from "./scope";

type AppliedBatch = {
  document: Document;
  inverse: readonly DocumentOp[];
  effects: NonNullable<SequencedBatch["effects"]>;
};

// The ids held in the paragraph an insertion addresses.
const paragraphIdentities = (
  document: Document,
  op: InsertContentOp,
): string[] =>
  identityKeysIn(
    storyParagraphs(storyBody(document, op.at.story)).find(
      ({ paragraph }) =>
        paragraph.paraId !== undefined &&
        idKey(paragraph.paraId) === idKey(op.at.blockId),
    )?.paragraph ?? [],
  );

/**
 * Whether an insertion cut an identified record: its paragraph holds an id
 * that neither it, the inserted content nor the insertion's own revision
 * held before.
 */
const cutIdentified = (
  before: Document,
  after: Document,
  op: InsertContentOp,
): boolean => {
  const known = new Set([
    ...paragraphIdentities(before, op),
    ...identityKeysIn(op.slice.content),
    ...(op.revision === undefined
      ? []
      : [slotKey({ space: IDENTITY_SPACES.REVISION, id: op.revision.id })]),
  ]);
  return paragraphIdentities(after, op).some((key) => !known.has(key));
};

/**
 * Where the run a direct format lies inside, and cuts, begins and ends: a
 * run of the paragraph's own, holding the whole range with more on at least
 * one side. Undefined for a tracked format, one carrying an inverse's
 * preconditions or joins, and one that cuts no such run.
 */
const cutRun = (
  document: Document,
  op: SetRunPropsOp,
): { from: number; to: number } | undefined => {
  if (
    op.revision !== undefined ||
    op.expected !== undefined ||
    op.joinStart !== undefined ||
    op.joinEnd !== undefined ||
    idKey(op.from.blockId) !== idKey(op.to.blockId)
  )
    return undefined;
  const paragraph = storyParagraphs(storyBody(document, op.from.story)).find(
    ({ paragraph: candidate }) =>
      candidate.paraId !== undefined &&
      idKey(candidate.paraId) === idKey(op.from.blockId),
  )?.paragraph;
  if (paragraph === undefined) return undefined;
  const own = new Set<unknown>(paragraph.content);
  const run = runGaps(paragraph.content).find(
    ({ node, before, after }) =>
      own.has(node) &&
      before.offset <= op.from.offset &&
      op.to.offset <= after.offset &&
      (before.offset < op.from.offset || op.to.offset < after.offset),
  );
  return run === undefined
    ? undefined
    : { from: run.before.offset, to: run.after.offset };
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
    let effect: NonNullable<SequencedBatch["effects"]>[number] = {
      type: NO_SEQUENCED_EFFECT,
    };
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
        effect = {
          type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
          firstLength: paragraphLength(paragraph),
        };
      }
    }
    if (op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS) {
      const cut = cutRun(current, op);
      effect = {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        ...(cut === undefined ? {} : { cutRun: cut }),
      };
    }
    const applied = applyDocumentOps(current, [op]);
    if (applied.isErr()) {
      return Result.err(applied.error);
    }
    if (op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT) {
      effect = {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        cutIdentified: cutIdentified(current, applied.value.document, op),
      };
    } else if (
      op.type === DOCUMENT_OP_TYPES.DELETE_BLOCKS ||
      op.type === DOCUMENT_OP_TYPES.RESOLVE_REVISION
    ) {
      const { modified, inserted, removed } = applied.value.touched;
      effect = {
        type: TOUCHED_BLOCKS_EFFECT,
        blockIds: [...new Set([...modified, ...inserted, ...removed])],
      };
    } else if (sequencedEffectType(op.type) === TOUCHED_BLOCKS_EFFECT) {
      effect = blockScopedEffect(
        current,
        applied.value.document,
        op,
        applied.value.touched,
      );
    }
    current = applied.value.document;
    inverses.push(applied.value.inverse);
    effects.push(effect);
  }
  return Result.ok({
    document: current,
    inverse: inverses.reverse().flat(),
    effects,
  });
};
