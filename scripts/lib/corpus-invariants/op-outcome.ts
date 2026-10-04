import { EnsureParaIdsError, ENSURE_PARA_IDS_REASONS } from "@stll/folio-core/docx/ensureParaIds";
import { failureFromError } from "../corpus-signature";
import {
  DECLARED_REFUSAL_REASONS,
  type CorpusInvariantOutcome,
  type DeclaredRefusalInvariant,
  type StageTimings,
} from "./contract";

type OpErrorOutcomeOptions = {
  invariant: DeclaredRefusalInvariant;
  error: unknown;
  timings: StageTimings;
};

/** Message text never grants a declared refusal. */
export const opErrorOutcome = ({
  invariant,
  error,
  timings,
}: OpErrorOutcomeOptions): CorpusInvariantOutcome => {
  if (
    error instanceof EnsureParaIdsError &&
    error.reason === ENSURE_PARA_IDS_REASONS.SIGNED_PACKAGE
  )
    return {
      status: "declared-refusal",
      refusal: { invariant, reason: DECLARED_REFUSAL_REASONS.SIGNED_PACKAGE },
      timings,
    };
  return { status: "evaluated", failures: [failureFromError(invariant, error)], timings };
};
