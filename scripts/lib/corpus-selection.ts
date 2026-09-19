/**
 * Narrow a gate run to named files.
 *
 * Fixing one signature needs the files that signature fires on, not the whole
 * corpus: a full census is the better part of an hour and a hundred megabytes
 * of work to re-prove what the baseline already records. `--only` runs the
 * handful under investigation, so a fix can be measured in seconds before and
 * after.
 *
 * A selection that matches nothing is an error, not an empty run. A mistyped
 * pattern would otherwise report "no failures" and read as a fix.
 *
 * This narrows a run; it never narrows the ratchet. A subset census cannot be
 * compared against a baseline measured over the whole corpus, so `--only`
 * refuses to run with `--check` and `write-baseline` never sees a subset.
 */

import { TaggedError } from "better-result";

import { fileIdOf } from "./corpus-census";

export class CorpusSelectionError extends TaggedError("CorpusSelectionError")<{
  message: string;
}> {}

/** A run's own spelling of a task, on the one file id everything else matches. */
export const corpusFileId = ({
  sourceId,
  relativePath,
}: {
  sourceId: string;
  relativePath: string;
}): string => fileIdOf({ sourceId, path: relativePath });

/**
 * `--only a/b.docx,c/d.docx`, repeated or comma-separated, or absent.
 *
 * `undefined` means the whole corpus, which is what every unfiltered caller
 * wants and what CI always runs.
 */
export const parseOnlySelection = (values: readonly string[]): readonly string[] | undefined => {
  if (values.length === 0) {
    return undefined;
  }
  const patterns = values
    .flatMap((value) => value.split(","))
    .map((pattern) => pattern.trim())
    .filter((pattern) => pattern.length > 0);
  if (patterns.length === 0) {
    throw new CorpusSelectionError({ message: "--only needs at least one file id" });
  }
  return [...new Set(patterns)];
};

/**
 * A pattern selects a file by its full id or by any part of it, so a whole
 * source (`apache-poi`) and a single file both work without a second flag.
 */
const matches = (id: string, pattern: string): boolean => id.includes(pattern);

export type SelectOnlyOptions<T> = {
  entries: readonly T[];
  patterns: readonly string[];
  idOf: (entry: T) => string;
};

export const selectOnly = <T>({ entries, patterns, idOf }: SelectOnlyOptions<T>): T[] => {
  const unmatched = new Set(patterns);
  const selected = entries.filter((entry) => {
    const id = idOf(entry);
    let kept = false;
    for (const pattern of patterns) {
      if (!matches(id, pattern)) {
        continue;
      }
      unmatched.delete(pattern);
      kept = true;
    }
    return kept;
  });
  if (unmatched.size > 0) {
    throw new CorpusSelectionError({
      message: `--only matched no corpus file: ${[...unmatched].sort().join(", ")}`,
    });
  }
  return selected;
};
