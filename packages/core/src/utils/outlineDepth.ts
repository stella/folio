import type { HeadingInfo } from "./headingCollector";

/** Maximum heading depth displayed in the document outline. */
export type OutlineDepth = 2 | 3 | "all";

export const DEFAULT_OUTLINE_DEPTH = 2;

/** Filter headings for the outline without changing their document order. */
export const filterHeadingsByDepth = (
  headings: readonly HeadingInfo[],
  depth: OutlineDepth,
): HeadingInfo[] => {
  if (depth === "all") return [...headings];
  return headings.filter((heading) => heading.level < depth);
};
