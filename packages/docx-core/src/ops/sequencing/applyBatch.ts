import { panic, Result } from "better-result";

import type { Document } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyBody, storyParagraphs } from "../blocks";
import { IDENTITY_SPACES, identityKeysIn, idKey, slotKey } from "../ids";
import { type Gap, leafSpans, spanningRecords } from "../leaves";
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

const formatParagraph = (document: Document, op: SetRunPropsOp) =>
  storyParagraphs(storyBody(document, op.from.story)).find(
    ({ paragraph }) =>
      paragraph.paraId !== undefined &&
      idKey(paragraph.paraId) === idKey(op.from.blockId),
  )?.paragraph;

/**
 * The records a direct format cuts at both of its ends, outermost first, by
 * their gaps before it: those holding leaves on both sides of its range
 * before the format and no longer after it. Undefined for a tracked format,
 * one carrying an inverse's preconditions or joins, and one that cuts none.
 */
const cutRecords = (
  before: Document,
  after: Document,
  op: SetRunPropsOp,
): { from: Gap; to: Gap }[] | undefined => {
  if (
    op.revision !== undefined ||
    op.expected !== undefined ||
    op.joinStart !== undefined ||
    op.joinEnd !== undefined ||
    idKey(op.from.blockId) !== idKey(op.to.blockId)
  )
    return undefined;
  const original = formatParagraph(before, op);
  const formatted = formatParagraph(after, op);
  if (original === undefined || formatted === undefined) return undefined;
  const gaps = [op.from, op.to].map(({ offset, zeroWidthBefore }): Gap => ({
    offset,
    zeroWidthBefore: zeroWidthBefore ?? 0,
  }));
  // Formatting cuts records from the innermost out and never one around
  // the runs, so the records still spanning the range lead the chain.
  const cut = spanningRecords(original.content, gaps, 0, 2).slice(
    spanningRecords(formatted.content, gaps, 0, 2).length,
  );
  if (cut.length === 0) return undefined;
  const leaves = leafSpans(original.content);
  return cut.map((record) => {
    const held = leaves.filter(
      ({ node, ancestors }) => node === record || ancestors.includes(record),
    );
    const first = held.at(0);
    const last = held.at(-1);
    if (first === undefined || last === undefined)
      return panic("A record holding leaves on both sides holds a leaf.");
    return { from: first.before, to: last.after };
  });
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
    const applied = applyDocumentOps(current, [op]);
    if (applied.isErr()) {
      return Result.err(applied.error);
    }
    if (op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT) {
      effect = {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        cutIdentified: cutIdentified(current, applied.value.document, op),
      };
    } else if (op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS) {
      const cut = cutRecords(current, applied.value.document, op);
      effect = {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        ...(cut === undefined ? {} : { cutRecords: cut }),
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
