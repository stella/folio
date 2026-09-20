/**
 * Price every package against the corpus, not against a wall-clock constant.
 *
 * Parse cost scales with bytes, so a fixed millisecond budget only ever says
 * which file is large. What a corpus can say instead is which file is *unlike*
 * the corpus, and saying that needs a model of what the corpus costs.
 *
 * Milliseconds per megabyte is not that model. It assumes cost passes through
 * the origin, and parsing does not: opening a package, reading its styles and
 * its theme, and building an empty document cost the same tens of milliseconds
 * whether the body is one paragraph or ten thousand. Dividing that fixed cost
 * by a small file's megabytes produces an enormous rate, so a median over the
 * whole corpus mixes two populations — small files priced mostly by overhead
 * and large files priced mostly by content — and lands between them, too high
 * to catch a quadratic path in a large file and too low to leave a small one
 * alone.
 *
 * The model here is affine, `parseMs ≈ intercept + slope · bytes`, fitted by
 * the median of pairwise slopes (Theil-Sen) so that the outliers being hunted
 * cannot drag the baseline towards themselves the way a least-squares fit
 * would. A file's verdict is then how many times its own prediction it cost,
 * which is dimensionless and no longer a function of its size. Peak resident
 * set gets its own fit and its own verdict, because memory amplification and
 * slowness are different defects: a package that parses at corpus speed while
 * leaving the worker holding a gigabyte is a finding this family would
 * otherwise never report.
 *
 * A fit over one run also divides out that run's load, but only the part of it
 * that held steady. A machine uniformly twice as slow moves every file and the
 * baseline together, so the ratio each file is judged by does not move. Load
 * that *drifts* is the one the fit cannot absorb: a run takes minutes, a
 * shared machine does not stay still for minutes, and a file parsed during a
 * spike is priced against a baseline set mostly by files parsed while things
 * were quiet. The same file then passes or fails between identical runs.
 *
 * Each file therefore carries what a fixed synthetic package cost beside it.
 * {@link normalizeForLoad} divides that drift out, and a run whose reference
 * blew its own budget reports `degraded` rather than a verdict, because at
 * that point the timings are noise and a number computed from noise is worse
 * than an admission.
 *
 * The verdict cannot be reached per file, and this module holds no per-file
 * runner: the gate already prices the one parse every invariant shares and
 * samples the resident set beside it. What lives here is the pure functions
 * that turn those observations into a verdict, {@link normalizeForLoad},
 * {@link fitCorpusCost} and {@link performanceFailures}, which the gate calls
 * once the whole run is in.
 */

import { type CorpusFailure, failureFromAssertion } from "../corpus-signature";
import { EXTENDED_CORPUS_INVARIANTS } from "./contract";

/** One file's parse cost, as the census stores it. */
export type CorpusFileCost = {
  bytes: number;
  parseMs: number;
  /**
   * What a fixed synthetic package cost to parse in the same process, beside
   * this file.
   *
   * The affine fit prices a file against the corpus, which removes its size
   * from the verdict but not the machine's load: the fit is over one run, so a
   * file parsed during a load spike is compared against a baseline set mostly
   * by files parsed while the machine was quiet, and the same file passes or
   * fails depending on what else the box was doing. The reference is the same
   * work every time, so what it costs is the machine, and dividing it out
   * leaves the code.
   *
   * Zero for a file that never reached the reference (one that failed to read,
   * for instance); {@link normalizeForLoad} declines a verdict on those.
   */
  referenceMs: number;
  /**
   * The worker process's resident set at the moment the file finished, not a
   * true peak: sampling a real peak needs either an allocator hook or a
   * polling thread, and the kernel is single-threaded by contract. The
   * end-of-file sample still answers the question the corpus asks, because the
   * finding is not "this file allocated a lot" but "this file left the worker
   * holding two gigabytes", which is exactly what the sample shows.
   */
  peakRssBytes: number;
};

/** The same record under the name the fit and verdict helpers read it by. */
export type PerformanceObservation = CorpusFileCost;

/**
 * What the run's median reference may cost before the family stops reporting.
 *
 * The reference is an empty package, a couple of milliseconds on a machine
 * doing nothing else and low tens of milliseconds on one under heavy
 * contention. This is set around a hundred times the unloaded figure, which is
 * far above any load the reading still resolves and low enough to catch a box
 * that has stopped making progress. Deliberately permissive: normalisation is
 * the mechanism, and this is only the backstop for the case where dividing by
 * the reference is dividing noise by noise. A backstop that fired on ordinary
 * contention would silence the family instead of steadying it.
 */
export const MAX_REFERENCE_MS = 250;

/** A cost that grows with bytes from a fixed floor. */
export type AffineCost = { intercept: number; slope: number };

/** What the corpus costs, in time and in retained memory, as a function of bytes. */
export type CorpusCostModel = { parseMs: AffineCost; peakRssBytes: AffineCost };

/**
 * Pairs the fit considers.
 *
 * Theil-Sen is defined over every pair, which is quadratic, and a full corpus
 * run is thousands of files. Striding the sorted observations keeps the fit
 * deterministic and linear in the cap while still crossing the whole size
 * range: every observation is paired with the one a fixed distance ahead of
 * it, for a set of distances spread across the corpus.
 */
const MAX_FIT_PAIRS = 200_000;

const median = (values: number[]): number | null => {
  if (values.length === 0) {
    return null;
  }
  values.sort((left, right) => left - right);
  const middle = values.length >> 1;
  if (values.length % 2 === 1) {
    // SAFETY: `middle` is below a length already proven non-zero.
    return values[middle] as number;
  }
  // SAFETY: an even non-zero length puts both `middle - 1` and `middle` in range.
  return ((values[middle - 1] as number) + (values[middle] as number)) / 2;
};

/**
 * The median pairwise slope, then the median residual as the intercept.
 *
 * Returns `null` when no pair has two distinct byte counts, which is the only
 * state in which the corpus cannot say what it costs. Ties are skipped rather
 * than treated as infinite slopes.
 */
const fitAffine = (points: readonly { x: number; y: number }[]): AffineCost | null => {
  const sorted = [...points].sort((left, right) => left.x - right.x);
  if (sorted.length < 2) {
    return null;
  }

  const slopes: number[] = [];
  const strideBudget = Math.max(1, Math.floor(MAX_FIT_PAIRS / sorted.length));
  const strideStep = Math.max(1, Math.floor(sorted.length / strideBudget));
  for (let stride = strideStep; stride < sorted.length; stride += strideStep) {
    for (let index = 0; index + stride < sorted.length; index += 1) {
      // SAFETY: both indices are below the length the loops are bounded by.
      const low = sorted[index] as { x: number; y: number };
      const high = sorted[index + stride] as { x: number; y: number };
      if (high.x !== low.x) {
        slopes.push((high.y - low.y) / (high.x - low.x));
      }
    }
  }

  const slope = median(slopes);
  if (slope === null) {
    return null;
  }
  const intercept = median(sorted.map(({ x, y }) => y - slope * x));
  if (intercept === null) {
    return null;
  }
  return { intercept, slope };
};

/**
 * The run's verdict about the machine it ran on.
 *
 * `degraded` is not a failure and not a pass: it is the family declining to
 * report, which is the one honest answer when the timings priced the box
 * rather than the code. It carries a reason because a degraded run that says
 * nothing about why is indistinguishable from a broken one.
 */
export type PerformanceLoad =
  | { readonly status: "measured"; readonly observations: readonly PerformanceObservation[] }
  | { readonly status: "degraded"; readonly reason: string };

/**
 * Divide the machine out of a run's timings, or decline the run.
 *
 * Each file's parse is scaled by how much slower the reference was beside it
 * than it was across the run, so a file measured during a spike is compared on
 * the same terms as one measured while the machine was quiet. Peak resident
 * set is left alone: it is a quantity of memory, not a rate, and it does not
 * grow because another process is busy.
 *
 * Scaling by the run's own median rather than by an absolute reference keeps
 * the output in milliseconds and keeps the fit's intercept meaningful; the
 * absolute check is {@link MAX_REFERENCE_MS}, applied to that median.
 */
export const normalizeForLoad = (
  observations: readonly PerformanceObservation[],
): PerformanceLoad => {
  const readings = observations
    .map(({ referenceMs }) => referenceMs)
    .filter((value) => value > 0 && Number.isFinite(value));
  if (readings.length === 0) {
    return {
      status: "degraded",
      reason: "no reference timing was recorded, so the run cannot price the machine",
    };
  }

  const reference = median(readings);
  if (reference === null || reference > MAX_REFERENCE_MS) {
    return {
      status: "degraded",
      reason: "the reference package exceeded its own budget, so the run measured load, not cost",
    };
  }

  return {
    status: "measured",
    observations: observations.map((observation) => {
      const usable = observation.referenceMs > 0 && Number.isFinite(observation.referenceMs);
      return {
        bytes: observation.bytes,
        parseMs: usable
          ? observation.parseMs * (reference / observation.referenceMs)
          : observation.parseMs,
        peakRssBytes: observation.peakRssBytes,
        referenceMs: observation.referenceMs,
      };
    }),
  };
};

/**
 * Fit both cost models over the whole census, or report that it cannot say.
 *
 * A zero-byte file carries no information about how cost grows with bytes and
 * is left out of the fit; {@link performanceFailures} declines a verdict on it
 * for the same reason.
 */
export const fitCorpusCost = (
  observations: readonly PerformanceObservation[],
): CorpusCostModel | null => {
  const sized = observations.filter(({ bytes }) => bytes > 0);
  const parseMs = fitAffine(sized.map(({ bytes, parseMs: y }) => ({ x: bytes, y })));
  const peakRssBytes = fitAffine(sized.map(({ bytes, peakRssBytes: y }) => ({ x: bytes, y })));
  if (!parseMs || !peakRssBytes) {
    return null;
  }
  return { parseMs, peakRssBytes };
};

/**
 * A robust fit can put the intercept at or below zero on a corpus whose small
 * files are cheap, and a prediction of zero makes every file infinitely above
 * it. Predictions are floored so that the verdict stays a statement about the
 * model rather than about its arithmetic.
 */
const MIN_PREDICTED_MS = 1;
const MIN_PREDICTED_RSS_BYTES = 1_048_576;

const predict = ({ intercept, slope }: AffineCost, bytes: number, floor: number): number =>
  Math.max(floor, intercept + slope * bytes);

const TEN_TIMES = 10;
const HUNDRED_TIMES = 100;

/**
 * Buckets, no numbers.
 *
 * A signature is the defect with every per-file particular erased, so the
 * message may not carry this file's milliseconds or the corpus's fitted
 * baseline. The bucket is the whole finding: a file at eleven times its
 * predicted cost and a file at five hundred times it are different defects,
 * and nothing between them is.
 */
const TEN_TIMES_MESSAGE = "parse cost exceeds ten times the corpus baseline for its size";
const HUNDRED_TIMES_MESSAGE =
  "parse cost exceeds one hundred times the corpus baseline for its size";
const RSS_MESSAGE = "peak resident set exceeds ten times the corpus baseline for its size";

/**
 * The outlier verdict for one observation against the fitted census.
 *
 * Time and memory are reported independently, so a file can be both. Every
 * threshold is exclusive: a file at exactly ten times its prediction is not
 * above ten times it, and a file at exactly one hundred times it falls in the
 * ten-times bucket. Inclusivity matters only for a corpus of near-identical
 * files, where an exclusive bound is the one that stays quiet.
 */
export const performanceFailures = (
  observation: PerformanceObservation,
  model: CorpusCostModel | null,
): CorpusFailure[] => {
  // A file with no reference reading was never priced against the machine, so
  // its milliseconds are not in the same unit as the fit and cannot be judged
  // by it. Its resident set still could be, but reporting half a verdict reads
  // as the other half having passed.
  if (!model || observation.bytes <= 0 || observation.referenceMs <= 0) {
    return [];
  }

  const failures: CorpusFailure[] = [];
  const predictedMs = predict(model.parseMs, observation.bytes, MIN_PREDICTED_MS);
  if (observation.parseMs > HUNDRED_TIMES * predictedMs) {
    failures.push(
      failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.performance, HUNDRED_TIMES_MESSAGE),
    );
  } else if (observation.parseMs > TEN_TIMES * predictedMs) {
    failures.push(failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.performance, TEN_TIMES_MESSAGE));
  }

  const predictedRss = predict(model.peakRssBytes, observation.bytes, MIN_PREDICTED_RSS_BYTES);
  if (observation.peakRssBytes > TEN_TIMES * predictedRss) {
    failures.push(failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.performance, RSS_MESSAGE));
  }

  return failures;
};
