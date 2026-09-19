/**
 * The invariants that run off the same parse as the gate's original four.
 *
 * Each one lives in its own module under `corpus-invariants/` and owns its own
 * baseline file, so adding an invariant never rewrites another's findings and
 * two branches working on different invariants do not meet in one file.
 *
 * Ordering is by cost, cheapest first. A file that exhausts its budget stops
 * here rather than in the middle of a stage, and the invariants it never
 * reached are reported as skipped instead of silently passing: a census that
 * counted a skipped invariant as a pass would shrink under load.
 */

import { Result } from "better-result";

import { classifyCorpusFile } from "./corpus-classify";
import {
  type CorpusInvariantInput,
  type ExtendedCorpusInvariant,
  EXTENDED_CORPUS_INVARIANTS,
  type StageTimings,
} from "./corpus-invariants/contract";
import { runEditLocalityInvariant } from "./corpus-invariants/edit-locality";
import { runEditorRoundTripInvariant } from "./corpus-invariants/editor-round-trip";
import { runKernelDifferentialInvariant } from "./corpus-invariants/kernel-differential";
import { runPipelineTotalityInvariant } from "./corpus-invariants/pipeline-totality";
import { runReserializeInvariant } from "./corpus-invariants/reserialize";
import { runSaveIdempotenceInvariant } from "./corpus-invariants/save-idempotence";
import { runSchemaValidityInvariant } from "./corpus-invariants/schema-validity";
import { type CorpusProducer, PRODUCER_FAMILIES, readCorpusProducer } from "./corpus-producer";
import {
  type CorpusFailure,
  type CorpusInvariant,
  failureFromAssertion,
  failureFromError,
} from "./corpus-signature";

/**
 * Every extended invariant except `performance`, which is not a per-file
 * verdict: whether a file is an outlier is a fact about the whole census, so
 * the gate decides it once the run is over.
 */
type RunnableInvariant = Exclude<
  ExtendedCorpusInvariant,
  typeof EXTENDED_CORPUS_INVARIANTS.performance
>;

const INVARIANT_RUNNERS = {
  [EXTENDED_CORPUS_INVARIANTS.schemaValidity]: runSchemaValidityInvariant,
  [EXTENDED_CORPUS_INVARIANTS.kernelDifferential]: runKernelDifferentialInvariant,
  [EXTENDED_CORPUS_INVARIANTS.editorRoundTrip]: runEditorRoundTripInvariant,
  [EXTENDED_CORPUS_INVARIANTS.reserialize]: runReserializeInvariant,
  [EXTENDED_CORPUS_INVARIANTS.saveIdempotence]: runSaveIdempotenceInvariant,
  [EXTENDED_CORPUS_INVARIANTS.editLocality]: runEditLocalityInvariant,
  [EXTENDED_CORPUS_INVARIANTS.pipelineTotality]: runPipelineTotalityInvariant,
} as const satisfies Record<
  RunnableInvariant,
  (input: CorpusInvariantInput) => Promise<{ failures: CorpusFailure[]; timings: StageTimings }>
>;

/** Cheapest first, so a budget that runs out costs the least evidence. */
const INVARIANT_ORDER = [
  EXTENDED_CORPUS_INVARIANTS.schemaValidity,
  EXTENDED_CORPUS_INVARIANTS.kernelDifferential,
  EXTENDED_CORPUS_INVARIANTS.editorRoundTrip,
  EXTENDED_CORPUS_INVARIANTS.reserialize,
  EXTENDED_CORPUS_INVARIANTS.saveIdempotence,
  EXTENDED_CORPUS_INVARIANTS.editLocality,
  EXTENDED_CORPUS_INVARIANTS.pipelineTotality,
] as const satisfies readonly RunnableInvariant[];

export type ExtendedChecksOptions = Omit<CorpusInvariantInput, "budgetMs"> & {
  /** Milliseconds one invariant may take on this file before the overrun is a finding. */
  invariantBudgetMs: number;
  /** Milliseconds all of them together may take before the rest are skipped. */
  fileBudgetMs: number;
  /**
   * Run only these invariants.
   *
   * The minimiser evaluates one file hundreds of times against one signature,
   * so paying for every other invariant on every candidate would make shrinking
   * a finding cost more than finding it. A census passes nothing and runs all.
   */
  only?: ReadonlySet<CorpusInvariant>;
};

export type ExtendedChecksResult = {
  producer: CorpusProducer;
  failures: CorpusFailure[];
  /** `<invariant>.<stage>` to milliseconds, for the performance census. */
  timings: StageTimings;
  /** The invariant the file budget ran out before, when it did. */
  truncatedAt?: string;
};

const BUDGET_INVARIANT = EXTENDED_CORPUS_INVARIANTS.performance;

export const runExtendedChecks = async ({
  bytes,
  buffer,
  parsed,
  documentPart,
  invariantBudgetMs,
  fileBudgetMs,
  only,
}: ExtendedChecksOptions): Promise<ExtendedChecksResult> => {
  const producer = await Result.tryPromise({
    try: () => readCorpusProducer({ bytes, documentPart }),
    catch: (cause: unknown) => cause,
  });
  const failures: CorpusFailure[] = [];
  const timings: StageTimings = {};
  const input: CorpusInvariantInput = {
    bytes,
    buffer,
    parsed,
    documentPart,
    budgetMs: invariantBudgetMs,
  };

  let spentMs = 0;
  let truncatedAt: string | undefined;
  for (const invariant of INVARIANT_ORDER) {
    if (only !== undefined && !only.has(invariant)) {
      continue;
    }
    if (spentMs > fileBudgetMs) {
      truncatedAt ??= invariant;
      failures.push(
        failureFromAssertion(
          BUDGET_INVARIANT,
          `the file time budget was exhausted before ${invariant} ran`,
        ),
      );
      continue;
    }
    const started = Bun.nanoseconds();
    // oxlint-disable-next-line no-await-in-loop -- the invariants share one parse and run in cost order
    const outcome = await Result.tryPromise({
      try: () => INVARIANT_RUNNERS[invariant](input),
      catch: (cause: unknown) => cause,
    });
    const elapsedMs = (Bun.nanoseconds() - started) / 1e6;
    spentMs += elapsedMs;
    if (outcome.isErr()) {
      failures.push(failureFromError(invariant, outcome.error));
      continue;
    }
    failures.push(...outcome.value.failures);
    for (const [stage, ms] of Object.entries(outcome.value.timings)) {
      timings[`${invariant}.${stage}`] = ms;
    }
    if (elapsedMs > invariantBudgetMs) {
      // The stage ran to completion, only slowly. That is a timing finding and
      // nothing more: the loop continues, so this file's later invariants are
      // still measured and its gating findings still count. Only the file
      // budget above, which skips what it has not reached, truncates.
      failures.push(
        failureFromAssertion(BUDGET_INVARIANT, `${invariant} exceeded its per-file time budget`),
      );
    }
  }

  return {
    producer: producer.isErr()
      ? { family: PRODUCER_FAMILIES.unknown, label: PRODUCER_FAMILIES.unknown }
      : producer.value,
    failures,
    timings,
    ...(truncatedAt === undefined ? {} : { truncatedAt }),
  };
};

/** Re-exported so `corpus-check.ts` names the classifier once. */
export { classifyCorpusFile };
