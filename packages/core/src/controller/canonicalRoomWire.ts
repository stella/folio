import { Result } from "better-result";
import * as v from "valibot";

import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  DOCUMENT_OP_TYPES,
  ID_LEASE_KINDS,
  ID_LEASE_REFUSAL_REASONS,
  type IdBlock,
  type IdLease,
  IdLeaseRefusal,
} from "@stll/docx-core/ops";

import type { CanonicalRoomSubmission } from "./canonicalRoomTransport";

/*
 * What a room's answers look like between two contexts (tabs, a server and
 * its clients): plain JSON. A rejection and a lease refusal are tagged errors
 * in memory and lose their class on the way, so they travel as their fields
 * and are rebuilt on arrival.
 */

const LEASE_KINDS = v.picklist(Object.values(ID_LEASE_KINDS));
const REVISION = v.pipe(v.number(), v.integer(), v.minValue(0));
const ID = v.pipe(v.number(), v.safeInteger(), v.minValue(1));

const REJECTION = v.object({
  reason: v.picklist(Object.values(BATCH_REJECTION_REASONS)),
  message: v.string(),
  opId: v.exactOptional(v.string()),
  opType: v.exactOptional(v.picklist(Object.values(DOCUMENT_OP_TYPES))),
  overType: v.exactOptional(v.picklist(Object.values(DOCUMENT_OP_TYPES))),
  idKind: v.exactOptional(LEASE_KINDS),
});

const SUBMISSION = v.variant("type", [
  v.object({ type: v.literal("ack"), opId: v.string(), rev: REVISION }),
  v.object({
    type: v.literal("reject"),
    opId: v.string(),
    reason: REJECTION,
    headRev: REVISION,
  }),
  v.object({ type: v.literal("disconnected"), opId: v.string() }),
]);

const BLOCK = v.object({ kind: LEASE_KINDS, start: ID, end: ID });

const REFUSAL = v.object({
  reason: v.picklist(Object.values(ID_LEASE_REFUSAL_REASONS)),
  message: v.string(),
  kind: v.exactOptional(LEASE_KINDS),
});

const GRANTED = v.variant("ok", [
  v.object({
    ok: v.literal(true),
    lease: v.object({ blocks: v.array(BLOCK) }),
  }),
  v.object({ ok: v.literal(false), refusal: REFUSAL }),
]);

const RENEWED = v.variant("ok", [
  v.object({ ok: v.literal(true), block: BLOCK }),
  v.object({ ok: v.literal(false), refusal: REFUSAL }),
]);

export type CanonicalWireSubmission = v.InferOutput<typeof SUBMISSION>;

export const encodeRoomSubmission = (
  outcome: CanonicalRoomSubmission,
): CanonicalWireSubmission => {
  switch (outcome.type) {
    case "ack":
      return { type: "ack", opId: outcome.opId, rev: outcome.rev };
    case "disconnected":
      return { type: "disconnected", opId: outcome.opId };
    case "reject": {
      const { reason } = outcome;
      return {
        type: "reject",
        opId: outcome.opId,
        headRev: outcome.headRev,
        reason: {
          reason: reason.reason,
          message: reason.message,
          ...(reason.opId === undefined ? {} : { opId: reason.opId }),
          ...(reason.opType === undefined ? {} : { opType: reason.opType }),
          ...(reason.overType === undefined
            ? {}
            : { overType: reason.overType }),
          ...(reason.idKind === undefined ? {} : { idKind: reason.idKind }),
        },
      };
    }
    default: {
      const exhaustive: never = outcome;
      return exhaustive;
    }
  }
};

/** A submission's answer as another context sent it; undefined when malformed. */
export const decodeRoomSubmission = (
  value: unknown,
): CanonicalRoomSubmission | undefined => {
  const parsed = v.safeParse(SUBMISSION, value);
  if (!parsed.success) return undefined;
  const outcome = parsed.output;
  switch (outcome.type) {
    case "ack":
    case "disconnected":
      return outcome;
    case "reject":
      return {
        type: "reject",
        opId: outcome.opId,
        headRev: outcome.headRev,
        reason: new BatchRejection(outcome.reason),
      };
    default: {
      const exhaustive: never = outcome;
      return exhaustive;
    }
  }
};

const refusalFields = (refusal: IdLeaseRefusal) => ({
  reason: refusal.reason,
  message: refusal.message,
  ...(refusal.kind === undefined ? {} : { kind: refusal.kind }),
});

export type CanonicalWireLeaseGrant = v.InferOutput<typeof GRANTED>;
export type CanonicalWireLeaseRenewal = v.InferOutput<typeof RENEWED>;

export const encodeLeaseGrant = (
  result: Result<IdLease, IdLeaseRefusal>,
): CanonicalWireLeaseGrant =>
  result.isOk()
    ? { ok: true, lease: { blocks: [...result.value.blocks] } }
    : { ok: false, refusal: refusalFields(result.error) };

export const decodeLeaseGrant = (
  value: unknown,
): Result<IdLease, IdLeaseRefusal> | undefined => {
  const parsed = v.safeParse(GRANTED, value);
  if (!parsed.success) return undefined;
  const answer = parsed.output;
  return answer.ok
    ? Result.ok(answer.lease)
    : Result.err(new IdLeaseRefusal(answer.refusal));
};

export const encodeLeaseRenewal = (
  result: Result<IdBlock, IdLeaseRefusal>,
): CanonicalWireLeaseRenewal =>
  result.isOk()
    ? { ok: true, block: result.value }
    : { ok: false, refusal: refusalFields(result.error) };

export const decodeLeaseRenewal = (
  value: unknown,
): Result<IdBlock, IdLeaseRefusal> | undefined => {
  const parsed = v.safeParse(RENEWED, value);
  if (!parsed.success) return undefined;
  const answer = parsed.output;
  return answer.ok
    ? Result.ok(answer.block)
    : Result.err(new IdLeaseRefusal(answer.refusal));
};

/**
 * A value as JSON carries it: what a structured clone or a Yjs document
 * would otherwise keep (undefined fields, class instances) is dropped here,
 * on the sending side, so every context reads the same thing.
 */
export const toWireJson = (value: unknown): unknown =>
  JSON.parse(JSON.stringify(value));
