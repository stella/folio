/**
 * What a refused comparison was refused for.
 *
 * `compareDocx` returns an error rather than an unverified redline, so on a
 * corpus of documents nobody authored for it the interesting number is not the
 * wall time but the refusal rate — and a refusal rate is only actionable split
 * by cause. A total says how much was refused; a bucket says which part of it
 * is one missing operation and which is a document the engine must not process
 * at all.
 *
 * The classification is a total function of the error union: a new error class
 * cannot land without a bucket, because the switch has no default.
 */

import type { CompareDocxError, CompareUnsupportedPart } from "@stll/folio-core/compare/types";

/**
 * Why one comparison produced no document.
 *
 * Split finely enough that a bucket names one fix. The round-trip buckets are
 * the self-check's verdict read structurally: the check compares two block
 * projections, so WHICH field of the projection diverged says which part of
 * the pipeline lost the difference.
 */
export const REFUSAL_BUCKETS = Object.freeze({
  "parse-base": "The base package could not be read into an editor model.",
  "parse-target": "The target package could not be read into an editor model.",
  "apply-refused": "The applier refused a derived operation.",
  "operation-limit": "The difference needs more operations than the engine generates.",
  serialize: "The redlined package could not be written back out.",
  "round-trip-invisible-structure":
    "The target's blocks are all there, in order, at coordinates the block model cannot reach.",
  "round-trip-block-count": "Accepting left a different number of blocks than the target has.",
  "round-trip-container": "Every block's text matched, but one landed in the wrong container.",
  "round-trip-style": "A block kept the base's paragraph style where the target changed it.",
  "round-trip-list-level": "A block kept the base's list level where the target changed it.",
  "round-trip-whitespace": "A block's text differs from the target's only in whitespace.",
  "round-trip-text": "A block's text does not match the target's.",
  "invalid-options": "The call's options were not usable.",
} as const);

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
 * One block of the self-check's projection. `projectStory` writes
 * `container|styleId|listLevel|text`, so the text is everything after the
 * third separator and may itself contain one.
 */
type ProjectedBlock = {
  container: string;
  styleId: string;
  listLevel: string;
  text: string;
};

const parseProjection = (entry: string): ProjectedBlock => {
  const first = entry.indexOf("|");
  const second = entry.indexOf("|", first + 1);
  const third = entry.indexOf("|", second + 1);
  if (first === -1 || second === -1 || third === -1) {
    return { container: "", styleId: "", listLevel: "", text: entry };
  }
  return {
    container: entry.slice(0, first),
    styleId: entry.slice(first + 1, second),
    listLevel: entry.slice(second + 1, third),
    text: entry.slice(third + 1),
  };
};

const containerKind = (container: string): "body" | "cell" =>
  container.startsWith("t") ? "cell" : "body";

const collapseWhitespace = (text: string): string => text.replace(/\s+/gu, " ").trim();

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

/**
 * Where two projections first diverge, and what diverged there. A projection
 * pair that matches everywhere but in length diverges at the shorter one's end.
 */
const firstDivergence = (
  accepted: readonly string[],
  target: readonly string[],
): { index: number; accepted: ProjectedBlock | null; target: ProjectedBlock | null } => {
  const shared = Math.min(accepted.length, target.length);
  for (let index = 0; index < shared; index++) {
    if (accepted[index] !== target[index]) {
      return {
        index,
        accepted: parseProjection(accepted[index] ?? ""),
        target: parseProjection(target[index] ?? ""),
      };
    }
  }
  return {
    index: shared,
    accepted: shared < accepted.length ? parseProjection(accepted[shared] ?? "") : null,
    target: shared < target.length ? parseProjection(target[shared] ?? "") : null,
  };
};

const CONTAINER_PATTERN = /^t(\d+)r(\d+)c(\d+)p(\d+)$/u;

/**
 * The same projection with every table coordinate renumbered by first
 * appearance, so it counts the blocks the model holds rather than the
 * paragraphs the package contains.
 *
 * The snapshot skips every empty textblock, so a cell holding a blank
 * paragraph reports its visible paragraph at `p1`, and a table whose first
 * rows are empty reports its visible rows starting at `r3`. No operation can
 * put a block at those coordinates, because no operation can create the empty
 * paragraphs that produce them. When two projections agree here and disagree
 * on the raw coordinates, the redline holds every block the target does, in
 * order, and the difference is one the block model cannot see — which is a
 * different finding from a redline that lost content, and is reported as one.
 */
const byVisibleOrdinal = (entries: readonly string[]): string[] => {
  const ordinals = new Map<string, number>();
  const counts = new Map<string, number>();
  const ordinalWithin = (scope: string, index: string): number => {
    const key = `${scope}:${index}`;
    const existing = ordinals.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const next = counts.get(scope) ?? 0;
    counts.set(scope, next + 1);
    ordinals.set(key, next);
    return next;
  };

  const paragraphCounts = new Map<string, number>();
  return entries.map((entry) => {
    const first = entry.indexOf("|");
    const match = CONTAINER_PATTERN.exec(entry.slice(0, first));
    if (!match) {
      return entry;
    }
    const [, table = "", row = "", cell = ""] = match;
    const tableOrdinal = ordinalWithin("t", table);
    const rowScope = `r${String(tableOrdinal)}`;
    const rowOrdinal = ordinalWithin(rowScope, row);
    const cellScope = `c${String(tableOrdinal)}.${String(rowOrdinal)}`;
    const cellOrdinal = ordinalWithin(cellScope, cell);
    // Separated, so cell 23 of scope `c0.1` is not cell 3 of scope `c0.12`.
    const cellKey = `${cellScope}:${cell}`;
    const paragraphOrdinal = paragraphCounts.get(cellKey) ?? 0;
    paragraphCounts.set(cellKey, paragraphOrdinal + 1);
    return (
      `t${String(tableOrdinal)}r${String(rowOrdinal)}c${String(cellOrdinal)}p${String(paragraphOrdinal)}` +
      entry.slice(first)
    );
  });
};

/** Element by element, never by joining: a block's text may hold the separator. */
const sameEntries = (left: readonly string[], right: readonly string[]): boolean =>
  left.length === right.length && left.every((entry, index) => entry === right[index]);

const classifyRoundTrip = (accepted: readonly string[], target: readonly string[]): Refusal => {
  if (sameEntries(byVisibleOrdinal(accepted), byVisibleOrdinal(target))) {
    return {
      bucket: "round-trip-invisible-structure",
      shape: `every block matches once table coordinates count visible blocks (${String(target.length)} blocks)`,
    };
  }
  const divergence = firstDivergence(accepted, target);
  const at = `at block ${String(divergence.index)}/${String(target.length)}`;
  if (divergence.accepted === null || divergence.target === null) {
    const side = accepted.length > target.length ? "more" : "fewer";
    return {
      bucket: "round-trip-block-count",
      shape: `accepting left ${side} blocks than the target has (${String(accepted.length)} vs ${String(target.length)}), diverging ${at}`,
    };
  }
  const { accepted: left, target: right } = divergence;
  const counts = `${String(accepted.length)} blocks accepted, ${String(target.length)} expected`;
  if (left.container !== right.container) {
    return {
      bucket: "round-trip-container",
      shape: `a block landed in a ${containerKind(left.container)} where the target has it in a ${containerKind(right.container)}, ${at} (${counts})`,
    };
  }
  if (left.text === right.text && left.styleId !== right.styleId) {
    return {
      bucket: "round-trip-style",
      shape: `the paragraph style did not move ${at} (${counts})`,
    };
  }
  if (left.text === right.text && left.listLevel !== right.listLevel) {
    return {
      bucket: "round-trip-list-level",
      shape: `the list level did not move ${at} (${counts})`,
    };
  }
  if (collapseWhitespace(left.text) === collapseWhitespace(right.text)) {
    return {
      bucket: "round-trip-whitespace",
      shape: `a block's text differs only in whitespace ${at} (${counts})`,
    };
  }
  if (accepted.length !== target.length) {
    return {
      bucket: "round-trip-block-count",
      shape: `accepting left ${String(accepted.length)} blocks against the target's ${String(target.length)}, first differing ${at}`,
    };
  }
  return {
    bucket: "round-trip-text",
    shape: `a ${containerKind(left.container)} block's text does not match ${at}, base-side length ${String(left.text.length)} against ${String(right.text.length)} (${counts})`,
  };
};

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
      return {
        bucket: "operation-limit",
        shape: `over ${String(error.limit)} operations`,
      };
    case "CompareDocxSerializeError":
      return { bucket: "serialize", shape: causeMessage(error.cause) };
    case "CompareDocxRoundTripError":
      return classifyRoundTrip(error.acceptedText, error.targetText);
    case "InvalidCompareDocxOptionsError":
      return { bucket: "invalid-options", shape: `option ${error.option}` };
    default: {
      const unreachable: never = error;
      throw new Error(`Unbucketed compare error: ${JSON.stringify(unreachable)}`);
    }
  }
};

/** One corpus pair's outcome, as the refusal report records it. */
export type PairOutcome =
  | {
      id: string;
      status: "produced";
      changes: number;
      unsupported: readonly CompareUnsupportedPart["reason"][];
    }
  | { id: string; status: "refused"; bucket: RefusalBucket; shape: string };

type BucketRow = { bucket: RefusalBucket; count: number; shapes: Map<string, number> };

/** The bucket table, largest first, with each bucket's commonest shape. */
export const summarizeRefusals = (
  outcomes: readonly PairOutcome[],
): { bucket: RefusalBucket; count: number; shape: string }[] => {
  const rows = new Map<RefusalBucket, BucketRow>();
  for (const outcome of outcomes) {
    if (outcome.status !== "refused") {
      continue;
    }
    const row = rows.get(outcome.bucket) ?? {
      bucket: outcome.bucket,
      count: 0,
      shapes: new Map<string, number>(),
    };
    row.count += 1;
    row.shapes.set(outcome.shape, (row.shapes.get(outcome.shape) ?? 0) + 1);
    rows.set(outcome.bucket, row);
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
