/**
 * Which licence tiers a run may use.
 *
 * Nothing is redistributed from any tier: that rule is enforced by the cache
 * living outside the repository, and it does not change here. Tiers govern
 * where a source may run, which is a separate question with a different answer
 * per source, and one the repository should not decide silently on the owner's
 * behalf.
 *
 * Selecting a tier changes which files a run sees, so it changes the corpus a
 * baseline was measured over. The digest below is taken over the selected files
 * alone, which is what makes a tier-1 baseline stay valid when a tier-2 source
 * is added to the manifest.
 */

import { TaggedError } from "better-result";

import {
  CORPUS_TIERS,
  type CorpusLock,
  type CorpusTier,
  DEFAULT_CORPUS_TIERS,
  corpusLockDigest,
} from "./corpus-manifest";

export class CorpusTierError extends TaggedError("CorpusTierError")<{ message: string }> {}

const TIER_VALUES: ReadonlySet<number> = new Set(Object.values(CORPUS_TIERS));

/** `--tiers 1,2`, or the CI default of tier 1 alone when the flag is absent. */
export const parseTierSelection = (value: string | undefined): CorpusTier[] => {
  if (value === undefined) {
    return [...DEFAULT_CORPUS_TIERS];
  }
  const requested = value
    .split(",")
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
  if (requested.length === 0) {
    throw new CorpusTierError({ message: "--tiers needs at least one tier" });
  }
  const tiers = new Set<CorpusTier>();
  for (const part of requested) {
    const tier = Number(part);
    if (!TIER_VALUES.has(tier)) {
      throw new CorpusTierError({ message: `--tiers expects 1, 2 or 3, got \`${part}\`` });
    }
    // SAFETY: the value was just proven a member of the tier union.
    tiers.add(tier as CorpusTier);
  }
  return [...tiers].sort((left, right) => left - right);
};

export const selectTiers = (lock: CorpusLock, tiers: readonly CorpusTier[]): CorpusLock => {
  const selected = new Set<number>(tiers);
  const sources = lock.sources.filter((source) => selected.has(source.tier));
  return {
    ...lock,
    sources,
    fileCount: sources.reduce((total, source) => total + source.files.length, 0),
    totalBytes: sources.reduce(
      (total, source) => total + source.files.reduce((bytes, file) => bytes + file.bytes, 0),
      0,
    ),
  };
};

/**
 * The digest of the corpus a run actually saw.
 *
 * Prefixed with the tier selection so that two runs over different tiers can
 * never compare against each other's baseline, even in the degenerate case
 * where one tier happens to contribute no files.
 */
export const tierScopedLockDigest = (lock: CorpusLock, tiers: readonly CorpusTier[]): string =>
  `t${tiers.join("")}-${corpusLockDigest(selectTiers(lock, tiers))}`;

export const describeTiers = (tiers: readonly CorpusTier[]): string => `tier ${tiers.join(", ")}`;
