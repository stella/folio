import { TaggedError } from "better-result";
import {
  failureRecord,
  type FailureMarker,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import type { BrowserInputAction } from "../visual/browserInputTrace";
import type { CanonicalFuzzObservation } from "./canonicalFuzzErrors";

export class CanonicalBrowserOracleError extends TaggedError("CanonicalBrowserOracleError")<{
  message: string;
  cause: unknown;
  observations: readonly CanonicalFuzzObservation[];
}> {}

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
  ...(failure instanceof CanonicalBrowserOracleError ? { observations: failure.observations } : {}),
});
