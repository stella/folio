import type { Result } from "better-result";

import type { Document } from "../../model/document";
import { validateOpsDocument } from "../contract";
import type { IdBlock, IdLease, IdLeaseKind } from "../idLease";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "../types";
import { captureDocumentOp, restoreDocumentOp } from "../wire";
import { applyBatch } from "./applyBatch";
import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  type DocumentBatch,
  type SequencedBatch,
  validateDocumentBatch,
} from "./envelope";
import {
  createIdLeaseTable,
  type IdLeaseRefusal,
  type IdLeaseState,
} from "./idLeases";
import { transformBatch } from "./transform";

const SUBMISSION_TYPES = { ACK: "ack", REJECT: "reject" } as const;

export type BatchAck = { type: "ack"; opId: string; rev: number };
export type BatchReject = {
  type: "reject";
  opId: string;
  reason: BatchRejection;
  headRev: number;
};
export type BatchSubmission = BatchAck | BatchReject;

export type SequencerOptions = {
  /** The room's lease table, as `idLeases` gave it before a restart. */
  idLeases?: IdLeaseState;
};

/**
 * Deterministic append-only authority; the caller supplies delivery and
 * storage. Every actor mints fresh ids only from the blocks the sequencer
 * leases it, and a batch introducing an id outside its actor's lease is
 * rejected, so concurrent actors never mint the same id.
 */
export const createSequencer = (
  document: Document,
  options: SequencerOptions = {},
): {
  submit: (batch: DocumentBatch) => BatchSubmission;
  /** The actor's remaining lease; the first call grants one block per kind. */
  grantIdLease: (actor: string) => Result<IdLease, IdLeaseRefusal>;
  /** One more block of `kind` for an actor running low. */
  renewIdLease: (
    actor: string,
    kind: IdLeaseKind,
  ) => Result<IdBlock, IdLeaseRefusal>;
  readonly document: Document;
  readonly headRev: number;
  readonly broadcasts: readonly SequencedBatch[];
  /** The lease table to persist with the journal. */
  readonly idLeases: IdLeaseState;
} => {
  const valid = validateOpsDocument(document);
  if (valid.isErr()) {
    throw valid.error;
  }
  const leaseTable = createIdLeaseTable(document, options.idLeases);
  if (leaseTable.isErr()) {
    throw leaseTable.error;
  }
  const leases = leaseTable.value;
  let current = document;
  const journal: SequencedBatch[] = [];
  // A replayed outcome answers only the actor whose batch first claimed the id.
  const results = new Map<
    string,
    { actor: unknown; outcome: BatchSubmission }
  >();

  const submit = (batch: DocumentBatch): BatchSubmission => {
    const previous = results.get(batch.opId);
    if (previous !== undefined) {
      if (previous.actor === batch.actor) return previous.outcome;
      return {
        type: SUBMISSION_TYPES.REJECT,
        opId: batch.opId,
        reason: new BatchRejection({
          reason: BATCH_REJECTION_REASONS.CONFLICT,
          message: "The operation id was already submitted by another actor.",
        }),
        headRev: journal.length,
      };
    }
    const reject = (reason: BatchRejection): BatchReject => {
      const outcome: BatchReject = {
        type: SUBMISSION_TYPES.REJECT,
        opId: batch.opId,
        reason,
        headRev: journal.length,
      };
      results.set(batch.opId, { actor: batch.actor, outcome });
      return outcome;
    };
    const validated = validateDocumentBatch(batch);
    if (validated.isErr()) {
      return reject(validated.error);
    }
    batch = validated.value;
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
      batch.ops.some(
        (op) =>
          op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK && op.newHalf === undefined,
      )
    ) {
      return reject(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.CONFLICT,
          message:
            "Stale split must state which half receives the new identity.",
        }),
      );
    }
    const transformed = transformBatch(
      validated.value,
      journal.slice(batch.baseRev),
    );
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
    const leased = leases.check(batch.actor, applied.value.document);
    if (leased.isErr()) {
      return reject(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.OUTSIDE_ID_LEASE,
          idKind: leased.error.kind,
          message: `The ${leased.error.kind} id ${leased.error.id} is outside the submitter's lease.`,
        }),
      );
    }
    const revision = journal.length + 1;
    const normalizedOps: DocumentOp[] = [];
    for (const [index, op] of transformed.value.ops.entries()) {
      const restored = restoreDocumentOp(op);
      if (restored.isErr()) {
        return reject(
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
            message: restored.error.message,
          }),
        );
      }
      const admitted = restored.value;
      const effect = applied.value.effects.at(index);
      normalizedOps.push(
        captureDocumentOp(
          admitted.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK &&
            effect?.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK
            ? { ...admitted, newHalf: effect.newHalf }
            : admitted,
        ),
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
    leases.record(batch.actor, current, leased.value);
    const ack: BatchAck = {
      type: SUBMISSION_TYPES.ACK,
      opId: batch.opId,
      rev: revision,
    };
    results.set(batch.opId, { actor: batch.actor, outcome: ack });
    return ack;
  };

  return {
    submit,
    grantIdLease: leases.grant,
    renewIdLease: leases.renew,
    get idLeases(): IdLeaseState {
      return leases.state;
    },
    get document(): Document {
      return current;
    },
    get headRev(): number {
      return journal.length;
    },
    get broadcasts(): readonly SequencedBatch[] {
      return journal.slice();
    },
  };
};
