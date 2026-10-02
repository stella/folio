/** Refusals are capability gaps, counted separately from applied cases and failures. */
export const createMissingOpBurndown = () => {
  const counts = new Map<string, number>();
  return {
    record: (kind: string) => counts.set(kind, (counts.get(kind) ?? 0) + 1),
    rows: () => [...counts].sort(([left], [right]) => left.localeCompare(right)),
    markdown: () =>
      [
        "| Missing operation | Refused intents |",
        "| --- | ---: |",
        ...[...counts]
          .sort(([left], [right]) => left.localeCompare(right))
          .map(([kind, count]) => `| missing op: ${kind} | ${count} |`),
      ].join("\n"),
  };
};
