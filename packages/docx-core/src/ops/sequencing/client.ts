import { Result, panic } from "better-result";

import type { Document } from "../../model/document";
import { validateOpsDocument } from "../contract";
import { equalForStaleness } from "../equality";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "../types";
import { captureDocumentOp, restoreDocumentOp } from "../wire";
import { applyBatch } from "./applyBatch";
import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  validateDocumentBatch,
  validateSequencedBatch,
  type DocumentBatch,
  type SequencedBatch,
} from "./envelope";
import type { BatchAck, BatchReject } from "./sequencer";
import { transformBatch } from "./transform";

const PENDING_STATUSES = {
  QUEUED: "queued",
  SUBMITTED: "submitted",
  ACKNOWLEDGED: "acknowledged",
} as const;

type PendingEntry = {
  batch: DocumentBatch;
  inverse: readonly DocumentOp[];
  effects: NonNullable<SequencedBatch["effects"]>;
  status: (typeof PENDING_STATUSES)[keyof typeof PENDING_STATUSES];
};

export type ClientNotice = { opId: string; reason: BatchRejection; ops: readonly DocumentOp[] };

/** Optimistic state with ordered authoritative replay and one causal submission in flight. */
export const createClient = (document: Document) => {
  const valid = validateOpsDocument(document);
  if (valid.isErr()) {
    throw valid.error;
  }
  let confirmed = document;
  let current = document;
  let headRev = 0;
  let entries: PendingEntry[] = [];
  let inFlight: string | undefined;
  const received = new Map<number, SequencedBatch>();
  const history = new Map<number, SequencedBatch>();
  const appliedIds = new Set<string>();
  const localIds = new Set<string>();
  const notices: ClientNotice[] = [];

  const drop = (entry: PendingEntry, reason: BatchRejection): void => {
    notices.push({ opId: entry.batch.opId, reason, ops: entry.batch.ops });
  };

  const replay = (): void => {
    current = confirmed;
    const surviving: PendingEntry[] = [];
    for (const [index, entry] of entries.entries()) {
      const applied = applyBatch(current, entry.batch.ops);
      if (applied.isErr()) {
        drop(
          entry,
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
            message: applied.error.message,
          }),
        );
        for (const rest of entries.slice(index + 1)) {
          drop(
            rest,
            new BatchRejection({
              reason: BATCH_REJECTION_REASONS.CONFLICT,
              message: "An earlier optimistic dependency could not be replayed.",
            }),
          );
        }
        break;
      }
      current = applied.value.document;
      surviving.push({ ...entry, inverse: applied.value.inverse, effects: applied.value.effects });
    }
    entries = surviving;
  };

  const enqueue = (batch: DocumentBatch): Result<DocumentBatch, BatchRejection> => {
    const validated = validateDocumentBatch(batch);
    if (validated.isErr()) {
      return validated;
    }
    batch = validated.value;
    if (batch.revision !== undefined || batch.baseRev !== headRev || localIds.has(batch.opId)) {
      return Result.err(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
          message: "Local batch must name the current revision and a fresh operation id.",
        }),
      );
    }
    const applied = applyBatch(current, batch.ops);
    if (applied.isErr()) {
      return Result.err(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
          message: applied.error.message,
        }),
      );
    }
    const normalizedOps: DocumentOp[] = [];
    for (const [index, op] of validated.value.ops.entries()) {
      const restored = restoreDocumentOp(op);
      if (restored.isErr()) {
        return Result.err(
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
    const normalized = { ...validated.value, ops: normalizedOps };
    entries.push({
      batch: normalized,
      inverse: applied.value.inverse,
      effects: applied.value.effects,
      status: PENDING_STATUSES.QUEUED,
    });
    localIds.add(batch.opId);
    current = applied.value.document;
    return Result.ok(normalized);
  };

  const nextSubmission = (): DocumentBatch | undefined => {
    if (inFlight !== undefined) {
      return undefined;
    }
    const first = entries.at(0);
    if (first === undefined) {
      return undefined;
    }
    const batch = { ...first.batch, baseRev: headRev };
    entries[0] = { ...first, batch, status: PENDING_STATUSES.SUBMITTED };
    inFlight = batch.opId;
    return batch;
  };

  const receiveAck = (ack: BatchAck): void => {
    entries = entries.map((entry) =>
      entry.batch.opId === ack.opId ? { ...entry, status: PENDING_STATUSES.ACKNOWLEDGED } : entry,
    );
    if (appliedIds.has(ack.opId) && inFlight === ack.opId) {
      inFlight = undefined;
    }
  };

  const receiveReject = (rejection: BatchReject): void => {
    const index = entries.findIndex(({ batch }) => batch.opId === rejection.opId);
    if (index < 0) {
      if (inFlight === rejection.opId) {
        inFlight = undefined;
      }
      return;
    }
    const entry = entries[index];
    if (entry === undefined) {
      return;
    }
    drop(entry, rejection.reason);
    const later = entries.slice(index + 1);
    entries.splice(index, 1);
    // Later batches were authored with the rejected edit present. Rebase them
    // over its inverse, refusing any whose dependency cannot be expressed.
    const originalPrefix = applyBatch(
      confirmed,
      entries
        .slice(0, index)
        .flatMap(({ batch }) => batch.ops)
        .concat(entry.batch.ops),
    );
    if (originalPrefix.isErr())
      return panic("Pending prefix must replay before a rejection is removed.");
    let originalCursor = originalPrefix.value.document;
    const inverseApplied = applyBatch(originalCursor, entry.inverse);
    if (inverseApplied.isErr())
      return panic("Recorded inverse must apply before later pending batches.");
    let inverse: SequencedBatch = {
      ...entry.batch,
      ops: entry.inverse,
      revision: headRev + 1,
      effects: inverseApplied.value.effects,
    };
    const retained = entries.slice(0, index);
    for (const [laterIndex, pending] of later.entries()) {
      const adjusted = transformBatch(pending.batch, [inverse]);
      const propagated = transformBatch(
        inverse,
        [{ ...pending.batch, revision: headRev + 1, effects: pending.effects }],
        { order: "before" },
      );
      if (adjusted.isErr()) {
        drop(pending, adjusted.error);
        for (const rest of later.slice(laterIndex + 1)) drop(rest, adjusted.error);
        break;
      }
      retained.push({ ...pending, batch: adjusted.value });
      if (propagated.isErr()) {
        for (const rest of later.slice(laterIndex + 1)) drop(rest, propagated.error);
        break;
      }
      const originalPending = applyBatch(originalCursor, pending.batch.ops);
      if (originalPending.isErr())
        return panic("Pending batch must replay in its original coordinate space.");
      originalCursor = originalPending.value.document;
      const captured = applyBatch(originalCursor, propagated.value.ops);
      if (captured.isErr()) {
        const reason = new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
          message: captured.error.message,
        });
        for (const rest of later.slice(laterIndex + 1)) drop(rest, reason);
        break;
      }
      inverse = { ...propagated.value, revision: headRev + 1, effects: captured.value.effects };
    }
    entries = retained;
    if (inFlight === rejection.opId) {
      inFlight = undefined;
    }
    replay();
  };

  const receiveBroadcast = (batch: SequencedBatch): Result<void, BatchRejection> => {
    const checked = validateSequencedBatch(batch);
    if (checked.isErr()) return checked;
    batch = checked.value;
    if (
      !Number.isSafeInteger(batch.revision) ||
      batch.revision < 1 ||
      batch.baseRev !== batch.revision - 1
    ) {
      return Result.err(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
          message: "Broadcast does not name an adjacent journal revision.",
        }),
      );
    }
    const existing = history.get(batch.revision) ?? received.get(batch.revision);
    if (existing !== undefined) {
      return equalForStaleness(existing, batch)
        ? Result.ok(undefined)
        : Result.err(
            new BatchRejection({
              reason: BATCH_REJECTION_REASONS.CONFLICT,
              message: "Broadcast conflicts with a previously delivered revision.",
            }),
          );
    }
    received.set(batch.revision, batch);
    while (received.has(headRev + 1)) {
      const next = received.get(headRev + 1);
      if (next === undefined) {
        break;
      }
      if (appliedIds.has(next.opId)) {
        return Result.err(
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.CONFLICT,
            message: "An operation id appears at two journal revisions.",
          }),
        );
      }
      const applied = applyBatch(confirmed, next.ops);
      if (applied.isErr()) {
        return Result.err(
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
            message: applied.error.message,
          }),
        );
      }
      if (localIds.has(next.opId)) {
        entries = entries.filter(({ batch: pendingBatch }) => pendingBatch.opId !== next.opId);
        if (inFlight === next.opId) {
          inFlight = undefined;
        }
      } else {
        let remote = next;
        let originalCursor = confirmed;
        const rebased: PendingEntry[] = [];
        for (const [index, pending] of entries.entries()) {
          const adjusted = transformBatch(pending.batch, [remote]);
          const propagated = transformBatch(
            remote,
            [{ ...pending.batch, revision: headRev + 1, effects: pending.effects }],
            { order: "before" },
          );
          if (adjusted.isErr()) {
            drop(pending, adjusted.error);
            for (const rest of entries.slice(index + 1)) drop(rest, adjusted.error);
            break;
          }
          rebased.push({ ...pending, batch: adjusted.value });
          if (propagated.isErr()) {
            for (const rest of entries.slice(index + 1)) drop(rest, propagated.error);
            break;
          }
          const originalPending = applyBatch(originalCursor, pending.batch.ops);
          if (originalPending.isErr())
            return panic("Pending batch must replay in its original coordinate space.");
          originalCursor = originalPending.value.document;
          const captured = applyBatch(originalCursor, propagated.value.ops);
          if (captured.isErr()) {
            const reason = new BatchRejection({
              reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
              message: captured.error.message,
            });
            for (const rest of entries.slice(index + 1)) drop(rest, reason);
            break;
          }
          remote = {
            ...propagated.value,
            revision: next.revision,
            effects: captured.value.effects,
          };
        }
        entries = rebased;
      }
      confirmed = applied.value.document;
      headRev = next.revision;
      appliedIds.add(next.opId);
      received.delete(headRev);
      history.set(headRev, next);
      replay();
    }
    return Result.ok(undefined);
  };

  return {
    enqueue,
    nextSubmission,
    receiveAck,
    receiveReject,
    receiveBroadcast,
    get document() {
      return current;
    },
    get headRev() {
      return headRev;
    },
    get pending(): readonly DocumentBatch[] {
      return entries
        .filter(({ status }) => status !== PENDING_STATUSES.ACKNOWLEDGED)
        .map(({ batch }) => batch);
    },
    get notices(): readonly ClientNotice[] {
      return notices.slice();
    },
  };
};
