import type { Run, RunPropertyChange, TextFormatting } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";

/** The single field mark lists every revision held by its invisible result runs. */
export const emptyFieldRunPropertyChanges = (runs: readonly Run[]): RunPropertyChange[] =>
  runs.flatMap((run) => run.propertyChanges ?? []);

const visibleFormatting = (runs: readonly Run[]): TextFormatting | undefined =>
  runs.find((run) => run.formatting !== undefined)?.formatting;

type ResolveEmptyFieldResultRunsOptions = {
  runs: readonly Run[];
  mode: "accept" | "reject";
  revisionIds: ReadonlySet<number> | null;
};

type ResolvedEmptyFieldResultRuns = {
  runs: Run[];
  visibleFormatting: TextFormatting | undefined;
  visibleFormattingChanged: boolean;
};

/** Resolve each authored run separately; a field mark carries their combined revisions. */
export const resolveEmptyFieldResultRuns = ({
  runs,
  mode,
  revisionIds,
}: ResolveEmptyFieldResultRunsOptions): ResolvedEmptyFieldResultRuns | null => {
  let changed = false;
  const resolved = runs.map((run): Run => {
    const matches = run.propertyChanges?.filter(
      (change) => revisionIds === null || revisionIds.has(change.info.id),
    );
    if (!matches?.length) return run;

    changed = true;
    const remaining = run.propertyChanges?.filter(
      (change) => revisionIds !== null && !revisionIds.has(change.info.id),
    );
    const formatting = mode === "reject" ? matches.at(0)?.previousFormatting : run.formatting;
    return {
      type: "run",
      content: run.content,
      ...(formatting !== undefined ? { formatting } : {}),
      ...(remaining?.length ? { propertyChanges: remaining } : {}),
      ...(run.preservedAttributes ? { preservedAttributes: run.preservedAttributes } : {}),
    };
  });
  if (!changed) return null;

  const before = visibleFormatting(runs);
  const after = visibleFormatting(resolved);
  return {
    runs: resolved,
    visibleFormatting: after,
    visibleFormattingChanged: canonicalJson(before) !== canonicalJson(after),
  };
};
