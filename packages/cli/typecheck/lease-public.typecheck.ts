import type { FlushOutcome, PublicFlushRequest, FlushRequest } from "../src/editor-lease";
import type { PublicErrorDetails } from "../src/errors";
import type { LockHolder, PublicLockHolder } from "../src/lock";
import type { FailureBody } from "../src/output";

type AssertFalse<T extends false> = T;

export type InternalHolderCannotBePublished = AssertFalse<
  LockHolder extends PublicLockHolder ? true : false
>;

export type InternalFlushRequestCannotBePublished = AssertFalse<
  FlushRequest extends PublicFlushRequest ? true : false
>;

export type InternalHolderCannotEnterErrorDetails = AssertFalse<
  { holder: LockHolder } extends PublicErrorDetails ? true : false
>;

export type InternalRequestCannotEnterErrorDetails = AssertFalse<
  { pending: FlushRequest[] } extends PublicErrorDetails ? true : false
>;

export type InternalHolderCannotEnterCliOrMcpEnvelope = AssertFalse<
  { holder: LockHolder } extends FailureBody["details"] ? true : false
>;

export type InternalHolderCannotEnterFlushStatus = AssertFalse<
  LockHolder extends Extract<FlushOutcome, { type: "timedOut" }>["holder"] ? true : false
>;
