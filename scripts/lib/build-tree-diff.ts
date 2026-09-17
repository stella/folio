// Comparison of two builds of the same package, for the reproducibility gate.
//
// Kept apart from the driver so the reporting can be tested against fabricated
// trees instead of two real builds.

import { panic } from "better-result";

import { renderReportDiff } from "./api-report-diff";

/** One build's emitted files, keyed by path relative to the output directory. */
export type BuildTree = ReadonlyMap<string, string>;

export type BuildTreeMismatch =
  | { kind: "only-in-baseline"; path: string }
  | { kind: "only-in-candidate"; path: string }
  | { kind: "content"; path: string; baseline: string; candidate: string };

export type BuildTreeComparison = {
  /** The first build's tree. */
  baseline: BuildTree;
  /** The second build's tree, from the same sources. */
  candidate: BuildTree;
};

/**
 * Where two builds of the same sources first disagree, or null when they agree.
 *
 * Paths are walked in sorted order so a rerun of a nondeterministic build names
 * the same file, rather than whichever one the filesystem happened to list
 * first.
 */
export const firstBuildTreeMismatch = ({
  baseline,
  candidate,
}: BuildTreeComparison): BuildTreeMismatch | null => {
  const paths = new Set([...baseline.keys(), ...candidate.keys()]);
  for (const path of [...paths].toSorted()) {
    const before = baseline.get(path);
    const after = candidate.get(path);
    if (before === undefined) return { kind: "only-in-candidate", path };
    if (after === undefined) return { kind: "only-in-baseline", path };
    if (before !== after) return { kind: "content", path, baseline: before, candidate: after };
  }
  return null;
};

export type RenderBuildTreeMismatchOptions = {
  mismatch: BuildTreeMismatch;
  /** Diff lines printed before the rest is summarized away. */
  maxDiffLines: number;
};

/** The differing path, and for a content difference the diff that proves it. */
export const renderBuildTreeMismatch = ({
  mismatch,
  maxDiffLines,
}: RenderBuildTreeMismatchOptions): string => {
  switch (mismatch.kind) {
    case "only-in-baseline":
      return `${mismatch.path}\n  emitted by the first build only`;
    case "only-in-candidate":
      return `${mismatch.path}\n  emitted by the second build only`;
    case "content":
      return `${mismatch.path}\n${renderReportDiff({
        baseline: mismatch.baseline,
        candidate: mismatch.candidate,
        maxLines: maxDiffLines,
      })}`;
    default:
      return panic(`Unhandled build-tree mismatch: ${JSON.stringify(mismatch)}`);
  }
};
