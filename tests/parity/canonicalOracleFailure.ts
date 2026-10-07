import { Result, TaggedError } from "better-result";
import {
  failureRecord,
  type FailureMarker,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import type { BrowserInputAction } from "../visual/browserInputTrace";
import type { CanonicalFuzzError, CanonicalFuzzObservation } from "./canonicalFuzzErrors";

type CanonicalErrorCapture = { status: "complete" } | { status: "unavailable"; message: string };

export class CanonicalBrowserOracleError extends TaggedError("CanonicalBrowserOracleError")<{
  message: string;
  cause: unknown;
  observations: readonly CanonicalFuzzObservation[];
  errorCapture: CanonicalErrorCapture;
}> {}

type CaptureCanonicalOracleFailureOptions = {
  cause: unknown;
  observations: CanonicalFuzzObservation[];
  drainErrors: () => Promise<CanonicalFuzzError[]>;
};

/** Collect reports even when snapshotting fails before the normal phase drain. */
export const captureCanonicalOracleFailure = async ({
  cause,
  observations,
  drainErrors,
}: CaptureCanonicalOracleFailureOptions) => {
  const captured = await Result.tryPromise({ try: drainErrors, catch: (error: unknown) => error });
  let errorCapture: CanonicalErrorCapture;
  if (captured.isOk()) {
    const observation = observations.at(-1);
    if (observation) observation.errors.push(...captured.value);
    else observations.push({ phase: { type: "load" }, errors: captured.value });
    errorCapture = { status: "complete" };
  } else {
    errorCapture = {
      status: "unavailable",
      message: captured.error instanceof Error ? captured.error.message : String(captured.error),
    };
  }
  return new CanonicalBrowserOracleError({
    message: cause instanceof Error ? cause.message : String(cause),
    cause,
    observations,
    errorCapture,
  });
};

type CanonicalOracleFailureRecordOptions = {
  marker: FailureMarker;
  failure: unknown;
  flow: readonly BrowserInputAction[];
};

/** Keep structured error evidence independent of the truncated assertion message. */
export const canonicalOracleFailureRecord = ({
  marker,
  failure,
  flow,
}: CanonicalOracleFailureRecordOptions) => ({
  ...failureRecord(marker, failure, { flow }),
  ...(failure instanceof CanonicalBrowserOracleError
    ? { observations: failure.observations, errorCapture: failure.errorCapture }
    : {}),
});
