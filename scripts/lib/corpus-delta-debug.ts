/**
 * Delta debugging over an ordered set of removable things.
 *
 * `ddmin` in the shape Zeller and Hildebrandt describe: keep splitting the set
 * that still reproduces into finer partitions, and whenever a complement still
 * reproduces, continue from it. The result is 1-minimal with respect to the
 * partitions tried: dropping any further chunk stops reproducing.
 *
 * `reproduces` is expensive here (it re-packs and re-parses a package), so the
 * search is bounded by a budget of calls and returns the best candidate found
 * when it runs out.
 */

export type DeltaDebugOptions<T> = {
  items: readonly T[];
  reproduces: (kept: readonly T[]) => Promise<boolean>;
  budget: number;
};

export type DeltaDebugResult<T> = {
  kept: readonly T[];
  evaluations: number;
  exhaustedBudget: boolean;
};

const partition = <T>(items: readonly T[], count: number): T[][] => {
  const chunks: T[][] = [];
  const size = items.length / count;
  for (let index = 0; index < count; index += 1) {
    const start = Math.round(index * size);
    const end = Math.round((index + 1) * size);
    if (end > start) {
      chunks.push(items.slice(start, end));
    }
  }
  return chunks;
};

export const deltaDebug = async <T>({
  items,
  reproduces,
  budget,
}: DeltaDebugOptions<T>): Promise<DeltaDebugResult<T>> => {
  let kept = [...items];
  let granularity = 2;
  let evaluations = 0;

  while (kept.length >= 2 && evaluations < budget) {
    const chunks = partition(kept, Math.min(granularity, kept.length));
    let reduced = false;
    for (const chunk of chunks) {
      if (evaluations >= budget) {
        break;
      }
      const excluded = new Set(chunk);
      const complement = kept.filter((item) => !excluded.has(item));
      evaluations += 1;
      // oxlint-disable-next-line no-await-in-loop -- each candidate is tested against the previous winner
      if (await reproduces(complement)) {
        kept = complement;
        granularity = Math.max(granularity - 1, 2);
        reduced = true;
        break;
      }
    }
    if (reduced) {
      continue;
    }
    if (granularity >= kept.length) {
      break;
    }
    granularity = Math.min(granularity * 2, kept.length);
  }

  return { kept, evaluations, exhaustedBudget: evaluations >= budget };
};
