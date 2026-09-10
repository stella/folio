type PositionedRun = {
  pmStart?: number;
  pmEnd?: number;
};

type PositionedPageBreak = {
  pmStart: number;
  pmEnd: number;
};

export type PageBreakRunPartition<TRun, TPageBreak> = {
  before: TRun[];
  pageBreak: TPageBreak;
};

export type PageBreakRunPartitionResult<TRun, TPageBreak> =
  | {
      type: "partitioned";
      partitions: PageBreakRunPartition<TRun, TPageBreak>[];
      remaining: TRun[];
      runVisits: number;
    }
  | {
      type: "overlap";
      pageBreak: TPageBreak;
      run: TRun;
      runVisits: number;
    };

/** Partition ordered runs with one monotonic cursor; each run is consumed once and peeked once. */
export const partitionRunsAtPageBreaks = <
  TRun extends PositionedRun,
  TPageBreak extends PositionedPageBreak,
>(
  runs: readonly TRun[],
  pageBreaks: readonly TPageBreak[],
): PageBreakRunPartitionResult<TRun, TPageBreak> => {
  const partitions: PageBreakRunPartition<TRun, TPageBreak>[] = [];
  let nextRunIndex = 0;
  let runVisits = 0;

  for (const pageBreak of pageBreaks) {
    const before: TRun[] = [];
    while (nextRunIndex < runs.length) {
      const run = runs[nextRunIndex];
      if (!run) {
        break;
      }
      runVisits += 1;
      if (run.pmEnd !== undefined && run.pmEnd <= pageBreak.pmStart) {
        before.push(run);
        nextRunIndex += 1;
        continue;
      }
      if (run.pmStart !== undefined && run.pmStart >= pageBreak.pmEnd) {
        break;
      }
      return { type: "overlap", pageBreak, run, runVisits };
    }
    partitions.push({ before, pageBreak });
  }

  return {
    type: "partitioned",
    partitions,
    remaining: runs.slice(nextRunIndex),
    runVisits,
  };
};
