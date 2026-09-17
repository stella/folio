// Rendering for the public-API snapshot gate's two failure modes.
//
// `api:check` used to name the entries that drifted and stop there, which reads
// as "the snapshot is stale" even when the generator was fed a stale build. Both
// failures now say what actually differs, so a red CI run can be read without
// reproducing the environment locally.

/** Longest common subsequence over lines, as the indices that pair up. */
const commonSubsequence = (
  left: readonly string[],
  right: readonly string[],
): [number, number][] => {
  const lengths: number[][] = Array.from({ length: left.length + 1 }, () =>
    Array.from({ length: right.length + 1 }, () => 0),
  );
  for (let i = left.length - 1; i >= 0; i--) {
    for (let j = right.length - 1; j >= 0; j--) {
      const row = lengths[i];
      const next = lengths[i + 1];
      if (!row || !next) continue;
      row[j] =
        left[i] === right[j] ? (next[j + 1] ?? 0) + 1 : Math.max(next[j] ?? 0, row[j + 1] ?? 0);
    }
  }

  const pairs: [number, number][] = [];
  let i = 0;
  let j = 0;
  while (i < left.length && j < right.length) {
    if (left[i] === right[j]) {
      pairs.push([i, j]);
      i++;
      j++;
      continue;
    }
    if ((lengths[i + 1]?.[j] ?? 0) >= (lengths[i]?.[j + 1] ?? 0)) i++;
    else j++;
  }
  return pairs;
};

export type RenderReportDiffOptions = {
  /** The snapshot in `api-reports/`, i.e. what the branch claims the surface is. */
  committed: string;
  /** What API Extractor just produced from the built declarations. */
  generated: string;
  /** Lines to print before truncating; the count of dropped lines is reported. */
  maxLines: number;
};

/**
 * A unified-style diff with `-` for the committed snapshot and `+` for the
 * freshly generated one. Deliberately not a real unified diff: no hunk headers
 * and no context, because the only question a reader has is which symbols moved.
 */
export const renderReportDiff = ({
  committed,
  generated,
  maxLines,
}: RenderReportDiffOptions): string => {
  const left = committed.split("\n");
  const right = generated.split("\n");
  const pairs = commonSubsequence(left, right);

  const lines: string[] = [];
  let i = 0;
  let j = 0;
  const flushTo = (leftEnd: number, rightEnd: number): void => {
    for (; i < leftEnd; i++) lines.push(`-${left[i] ?? ""}`);
    for (; j < rightEnd; j++) lines.push(`+${right[j] ?? ""}`);
  };
  for (const [leftIndex, rightIndex] of pairs) {
    flushTo(leftIndex, rightIndex);
    i++;
    j++;
  }
  flushTo(left.length, right.length);

  if (lines.length <= maxLines) return lines.join("\n");
  const dropped = lines.length - maxLines;
  return [...lines.slice(0, maxLines), `… ${dropped} more diff line(s) truncated`].join("\n");
};

export type StaleBuildInput = {
  /** Modification time of the built declaration the report was extracted from. */
  declarationModifiedMs: number;
  /** Modification time of the newest source file in the same package. */
  newestSourceModifiedMs: number;
};

/**
 * Whether a declaration predates the sources it claims to describe.
 *
 * A stale `dist` makes the gate compare a snapshot of the branch against a
 * report of some earlier tree, which surfaces as drift that no `api:update` can
 * settle. Equal timestamps pass: a build that finished in the same millisecond
 * as the last source write is not evidence of staleness.
 */
export const isStaleDeclaration = ({
  declarationModifiedMs,
  newestSourceModifiedMs,
}: StaleBuildInput): boolean => declarationModifiedMs < newestSourceModifiedMs;
