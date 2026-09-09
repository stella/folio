/**
 * What a comparison could not prove, and what it refused outright.
 *
 * `compareDocx` verifies its own work: accepting the generated revisions must
 * reproduce the target, rejecting them must reproduce the base. It refuses by
 * default when it cannot prove that, and returns its best attempt plus the
 * list of failing invariants when asked with `onUnverified: "emit"`.
 *
 * This module reports both, over a corpus of documents nobody authored for the
 * engine: the refusal rate under the strict default, and the verified share
 * under the best-effort option. Both are only actionable split by cause, and
 * the causes come from the engine's own typed verdict rather than from a
 * second classification here — a copy of that judgement would drift from the
 * one that ships.
 */

import type { CompareDocxError, CompareUnsupportedPart } from "@stll/folio-core/compare/types";
import type {
  CompareVerification,
  CompareVerificationCause,
  CompareVerificationFailure,
} from "@stll/folio-core/compare/verification";

/**
 * A bucket per verification cause, so a cause the engine gains cannot land in
 * the report without a description. The prefix is what tells a reader that the
 * self-check produced the refusal rather than the parser or the serializer.
 */
type RoundTripBuckets = { [Cause in CompareVerificationCause as `round-trip-${Cause}`]: string };

/**
 * Why one comparison produced nothing under the strict default.
 *
 * The round-trip buckets are the engine's verification causes, prefixed so a
 * reader can tell a refusal the self-check produced from one the parser or the
 * serializer did, and checked against them so neither side can gain one alone.
 */
export const REFUSAL_BUCKETS = Object.freeze({
  "parse-base": "The base package could not be read into an editor model.",
  "parse-target": "The target package could not be read into an editor model.",
  "apply-refused": "The applier refused a derived operation.",
  "operation-limit": "The difference needs more operations than the engine generates.",
  serialize: "The redlined package could not be written back out.",
  "final-paragraph-mark":
    "A container's final paragraph mark carried a deletion no consumer could resolve.",
  "invalid-options": "The call's options were not usable.",
  "round-trip-invisible-structure":
    "Every block is there, in order, at coordinates the block model cannot reach.",
  "round-trip-block-count": "The result holds a different number of blocks than expected.",
  "round-trip-container": "Every block's text matched, but one sits in the wrong container.",
  "round-trip-inline-structure":
    "A block's explicit inline structure differs from the expected result.",
  "round-trip-table-geometry":
    "Every block is where it should be, and a table's own properties are not.",
  "round-trip-style": "A block kept a paragraph style the other side changed.",
  "round-trip-list-level": "A block kept a list level the other side changed.",
  "round-trip-alignment": "A block kept direct paragraph alignment the other side changed.",
  "round-trip-spacing": "A block kept direct paragraph spacing the other side changed.",
  "round-trip-inline-formatting": "A planned formatting change did not round-trip.",
  "round-trip-whitespace": "A block's text differs only in whitespace.",
  "round-trip-text": "A block's text does not match.",
} as const satisfies Record<string, string> & RoundTripBuckets);

export type RefusalBucket = keyof typeof REFUSAL_BUCKETS;

/** One refusal, bucketed, with a description that carries no document text. */
export type Refusal = {
  bucket: RefusalBucket;
  /**
   * The refusal's shape: structural facts only — counts, offsets, container
   * kinds, normalized cause names. Never a phrase of either document: this
   * string has to be safe to quote anywhere.
   */
  shape: string;
};

/**
 * An error message with its variable parts removed, so two failures of one
 * cause land in one bucket. Quoted fragments and numbers are the parts that
 * carry a document's own words and sizes.
 */
const messageShape = (message: string): string =>
  message
    .replace(/"[^"]*"/gu, '"…"')
    .replace(/'[^']*'/gu, "'…'")
    .replace(/\(paragraph [^)]*\)/gu, "")
    .replace(/\d+/gu, "N")
    .replace(/\s+/gu, " ")
    .trim()
    .slice(0, 180);

const causeMessage = (cause: unknown): string =>
  cause instanceof Error ? `${cause.name}: ${messageShape(cause.message)}` : "non-Error cause";

/** The bucket one verification cause belongs to. */
export const bucketOfCause = (cause: CompareVerificationFailure["cause"]): RefusalBucket =>
  `round-trip-${cause}`;

/** Bucket one refusal. Total over the error union: no default branch. */
export const classifyRefusal = (error: CompareDocxError): Refusal => {
  switch (error._tag) {
    case "CompareDocxParseError":
      return {
        bucket: error.side === "base" ? "parse-base" : "parse-target",
        shape: causeMessage(error.cause),
      };
    case "CompareDocxApplyError": {
      const reasons = [...new Set(error.skipped.map(({ reason }) => reason))].toSorted();
      return {
        bucket: "apply-refused",
        shape: `${String(error.skipped.length)} operation(s) refused: ${reasons.join(", ")}`,
      };
    }
    case "CompareDocxOperationLimitError":
      return { bucket: "operation-limit", shape: `over ${String(error.limit)} operations` };
    case "CompareDocxSerializeError":
      return { bucket: "serialize", shape: causeMessage(error.cause) };
    case "CompareDocxFinalParagraphMarkError":
      return {
        bucket: "final-paragraph-mark",
        shape: `${String(error.revisions.length)} container(s), first ${
          error.revisions.at(0)?.container ?? "unknown"
        }`,
      };
    case "CompareDocxRoundTripError":
      return {
        bucket: bucketOfCause(error.cause),
        shape: `${error.invariant}: ${error.failures.at(0)?.detail ?? "no detail"}`,
      };
    case "InvalidCompareDocxOptionsError":
      return { bucket: "invalid-options", shape: `option ${error.option}` };
    default: {
      const unreachable: never = error;
      throw new Error(`Unbucketed compare error: ${JSON.stringify(unreachable)}`);
    }
  }
};

/** One corpus pair's outcome under both modes. */
export type PairOutcome = {
  id: string;
  /** The strict default: a redline, or nothing. */
  strict: { status: "produced" } | ({ status: "refused" } & Refusal);
  /** Best effort: a redline either way, and whether it was proven. */
  bestEffort:
    | {
        status: "verified";
        changes: number;
        unsupported: readonly CompareUnsupportedPart["reason"][];
      }
    | {
        status: "unverified";
        changes: number;
        buckets: readonly RefusalBucket[];
        shapes: readonly string[];
      }
    | { status: "failed"; bucket: RefusalBucket; shape: string };
};

/** The best-effort side of one comparison, from its verification verdict. */
export const summarizeVerification = (
  verification: CompareVerification,
): { buckets: RefusalBucket[]; shapes: string[] } => {
  if (verification.status === "verified") {
    return { buckets: [], shapes: [] };
  }
  return {
    buckets: verification.failures.map(({ cause }) => bucketOfCause(cause)),
    shapes: verification.failures.map(({ invariant, detail }) => `${invariant}: ${detail}`),
  };
};

type BucketRow = { bucket: RefusalBucket; count: number; shapes: Map<string, number> };

const tally = (entries: readonly { bucket: RefusalBucket; shape: string }[]) => {
  const rows = new Map<RefusalBucket, BucketRow>();
  for (const { bucket, shape } of entries) {
    const row = rows.get(bucket) ?? { bucket, count: 0, shapes: new Map<string, number>() };
    row.count += 1;
    row.shapes.set(shape, (row.shapes.get(shape) ?? 0) + 1);
    rows.set(bucket, row);
  }
  return [...rows.values()]
    .toSorted((left, right) => right.count - left.count || left.bucket.localeCompare(right.bucket))
    .map(({ bucket, count, shapes }) => ({
      bucket,
      count,
      shape:
        [...shapes.entries()].toSorted((left, right) => right[1] - left[1])[0]?.[0] ?? "no shape",
    }));
};

/** The strict mode's refusals, largest bucket first. */
export const summarizeRefusals = (outcomes: readonly PairOutcome[]) =>
  tally(
    outcomes.flatMap(({ strict }) =>
      strict.status === "refused" ? [{ bucket: strict.bucket, shape: strict.shape }] : [],
    ),
  );

/**
 * The best-effort mode's unproven invariants, largest bucket first. One
 * document can contribute more than one: a redline that loses a difference
 * usually fails both directions of the round trip.
 */
export const summarizeUnverified = (outcomes: readonly PairOutcome[]) =>
  tally(
    outcomes.flatMap(({ bestEffort }) => {
      if (bestEffort.status === "failed") {
        return [{ bucket: bestEffort.bucket, shape: bestEffort.shape }];
      }
      if (bestEffort.status === "verified") {
        return [];
      }
      return bestEffort.buckets.map((bucket, index) => ({
        bucket,
        shape: bestEffort.shapes[index] ?? "no shape",
      }));
    }),
  );
