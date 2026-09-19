/**
 * Price every package against the corpus, not against a wall-clock constant.
 *
 * Parse cost scales with bytes, so a fixed millisecond budget only ever says
 * which files are large. What a corpus can say instead is which file is
 * *unlike* the corpus: a package whose milliseconds per megabyte are an order
 * of magnitude above the census median is running a quadratic path the other
 * thousands of files avoid, and that is a defect however fast the machine is.
 *
 * The verdict therefore cannot be reached per file, and this module holds no
 * per-file runner: the gate already prices the one parse every invariant shares
 * and samples the resident set beside it. What lives here is the pair of pure
 * functions that turn those observations into a verdict, `medianMsPerMegabyte`
 * and `performanceFailures`, which the gate calls once the whole run is in.
 */

import { type CorpusFailure, failureFromAssertion } from "../corpus-signature";
import { EXTENDED_CORPUS_INVARIANTS } from "./contract";

/** One file's parse cost, as the census stores it. */
export type CorpusFileCost = {
  bytes: number;
  parseMs: number;
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

/** The same record under the name the median and verdict helpers read it by. */
export type PerformanceObservation = CorpusFileCost;

const BYTES_PER_MEGABYTE = 1_048_576;

/** A file with no bytes has no cost per megabyte, so it cannot be an outlier. */
const msPerMegabyte = ({ bytes, parseMs }: PerformanceObservation): number | null =>
  bytes > 0 ? parseMs / (bytes / BYTES_PER_MEGABYTE) : null;

/**
 * The corpus's middle cost per megabyte, or `0` when the census cannot say.
 *
 * `0` covers both an empty census and one whose every file parsed in
 * unmeasurable time. Neither supports a multiple-of-the-median verdict, and
 * `performanceFailures` declines to render one rather than treating every file
 * as infinitely above zero.
 */
export const medianMsPerMegabyte = (observations: readonly PerformanceObservation[]): number => {
  const rates: number[] = [];
  for (const observation of observations) {
    const rate = msPerMegabyte(observation);
    if (rate !== null) {
      rates.push(rate);
    }
  }
  if (rates.length === 0) {
    return 0;
  }
  rates.sort((left, right) => left - right);
  const middle = rates.length >> 1;
  if (rates.length % 2 === 1) {
    // SAFETY: `middle` is below a length already proven non-zero.
    return rates[middle] as number;
  }
  // SAFETY: an even non-zero length puts both `middle - 1` and `middle` in range.
  return ((rates[middle - 1] as number) + (rates[middle] as number)) / 2;
};

const TEN_TIMES_MEDIAN = 10;
const HUNDRED_TIMES_MEDIAN = 100;

/**
 * Two buckets, no numbers.
 *
 * A signature is the defect with every per-file particular erased, so the
 * message may not carry this file's milliseconds or the census's median. The
 * bucket is the whole finding: a file at eleven times the median and a file at
 * five hundred times it are different defects, and nothing between them is.
 */
const TEN_TIMES_MESSAGE = "parse cost exceeds ten times the corpus median per megabyte";
const HUNDRED_TIMES_MESSAGE = "parse cost exceeds one hundred times the corpus median per megabyte";

/**
 * The outlier verdict for one observation against the whole census.
 *
 * Both thresholds are exclusive: a file at exactly ten times the median is not
 * above ten times it, and a file at exactly one hundred times the median falls
 * in the ten-times bucket. Inclusivity matters only for a corpus of near-
 * identical files, where an exclusive bound is the one that stays quiet.
 */
export const performanceFailures = (
  observation: PerformanceObservation,
  median: number,
): CorpusFailure[] => {
  const rate = msPerMegabyte(observation);
  if (rate === null || median <= 0) {
    return [];
  }
  if (rate > HUNDRED_TIMES_MEDIAN * median) {
    return [failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.performance, HUNDRED_TIMES_MESSAGE)];
  }
  if (rate > TEN_TIMES_MEDIAN * median) {
    return [failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.performance, TEN_TIMES_MESSAGE)];
  }
  return [];
};
