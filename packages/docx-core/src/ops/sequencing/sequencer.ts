import type { Document } from "../../model/document";
import { validateOpsDocument } from "../contract";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "../types";
import { applyBatch } from "./applyBatch";
import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  validateDocumentBatch,
  type DocumentBatch,
  type SequencedBatch,
} from "./envelope";
import { transformBatch } from "./transform";

const SUBMISSION_TYPES = { ACK: "ack", REJECT: "reject" } as const;

export type BatchAck = { type: "ack"; opId: string; rev: number };
export type BatchReject = { type: "reject"; opId: string; reason: BatchRejection; headRev: number };
export type BatchSubmission = BatchAck | BatchReject;

/** Deterministic append-only authority; the caller supplies delivery and storage. */
export const createSequencer = (document: Document) => {
  const valid = validateOpsDocument(document);
  if (valid.isErr()) {
    throw valid.error;
  }
  let current = document;
  const journal: SequencedBatch[] = [];
  const results = new Map<string, BatchSubmission>();

  const submit = (batch: DocumentBatch): BatchSubmission => {
    const previous = results.get(batch.opId);
    if (previous !== undefined) {
      return previous;
    }
    const reject = (reason: BatchRejection): BatchReject => {
      const outcome: BatchReject = {
        type: SUBMISSION_TYPES.REJECT,
        opId: batch.opId,
        reason,
        headRev: journal.length,
      };
      results.set(batch.opId, outcome);
      return outcome;
    };
    const validated = validateDocumentBatch(batch);
    if (validated.isErr()) {
      return reject(validated.error);
    }
    if (batch.revision !== undefined || batch.baseRev > journal.length) {
      return reject(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.STALE_BASE,
          message: "Submission does not name an available base revision.",
        }),
      );
    }
    if (
      batch.baseRev < journal.length &&
      batch.ops.some((op) => op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK && op.newHalf === undefined)
    ) {
      return reject(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.CONFLICT,
          message: "Stale split must state which half receives the new identity.",
        }),
      );
    }
    const transformed = transformBatch(validated.value, journal.slice(batch.baseRev));
    if (transformed.isErr()) {
      return reject(transformed.error);
    }
    const applied = applyBatch(current, transformed.value.ops);
    if (applied.isErr()) {
      return reject(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
          message: applied.error.message,
        }),
      );
    }
    const revision = journal.length + 1;
    const normalizedOps: DocumentOp[] = [];
    for (const [index, op] of transformed.value.ops.entries()) {
      const effect = applied.value.effects.at(index);
      normalizedOps.push(
        op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK && effect?.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK
          ? { ...op, newHalf: effect.newHalf }
          : op,
      );
    }
    const sequenced: SequencedBatch = {
      ...transformed.value,
      ops: normalizedOps,
      baseRev: journal.length,
      revision,
      effects: applied.value.effects,
    };
    journal.push(sequenced);
    current = applied.value.document;
    const ack: BatchAck = { type: SUBMISSION_TYPES.ACK, opId: batch.opId, rev: revision };
    results.set(batch.opId, ack);
    return ack;
  };

  return {
    submit,
    get document() {
      return current;
    },
    get headRev() {
      return journal.length;
    },
    get broadcasts(): readonly SequencedBatch[] {
      return journal.slice();
    },
  };
};
