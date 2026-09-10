export type BookmarkBoundaryOccurrence = {
  id: number;
  type: "start" | "end";
  position: number;
};

type BoundarySummary = {
  startPositions: number[];
  endPositions: number[];
};

export const findInvalidBookmarkBoundaryIds = (
  occurrences: readonly BookmarkBoundaryOccurrence[],
  reservedIds: ReadonlySet<number> = new Set(),
): ReadonlySet<number> => {
  const summaries = new Map<number, BoundarySummary>();

  for (const occurrence of occurrences) {
    const summary = summaries.get(occurrence.id) ?? {
      startPositions: [],
      endPositions: [],
    };
    if (occurrence.type === "start") {
      summary.startPositions.push(occurrence.position);
    } else {
      summary.endPositions.push(occurrence.position);
    }
    summaries.set(occurrence.id, summary);
  }

  const invalidIds = new Set<number>();
  for (const [id, summary] of summaries) {
    const start = summary.startPositions.at(0);
    const end = summary.endPositions.at(0);
    if (
      reservedIds.has(id) ||
      summary.startPositions.length !== 1 ||
      summary.endPositions.length !== 1 ||
      start === undefined ||
      end === undefined ||
      start >= end
    ) {
      invalidIds.add(id);
    }
  }
  return invalidIds;
};
