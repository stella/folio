import { findPageIndexContainingPmPos } from "../layout-engine/pmPageIndex";
import type { Layout } from "../layout-engine/types";
import type { HeadingInfo } from "./headingCollector";

const sameHeading = (a: HeadingInfo, b: HeadingInfo): boolean =>
  Object.values({
    id: a.id === b.id,
    text: a.text === b.text,
    level: a.level === b.level,
    pmPos: a.pmPos === b.pmPos,
    pageNumber: (a.pageNumber ?? null) === (b.pageNumber ?? null),
  } satisfies Record<keyof HeadingInfo, boolean>).every(Boolean);

const layoutPageNumber = (layout: Layout, pmPos: number): number | null => {
  const pageIndex = findPageIndexContainingPmPos(layout, pmPos);
  return pageIndex === null ? null : (layout.pages[pageIndex]?.number ?? null);
};

type ResolveOutlineHeadingsOptions = {
  /** Headings collected from the document `layout` was computed from. */
  collected: readonly HeadingInfo[];
  layout: Layout | null;
  previous: HeadingInfo[];
};

/**
 * The outline's headings with page numbers from `layout`.
 *
 * A heading the layout does not place keeps the page number it last had under
 * the same id, so the outline never shows a transient blank while a layout is
 * pending. An unchanged entry is returned as its previous object and an
 * unchanged list as the previous array, so consumers can skip work by identity.
 */
export const resolveOutlineHeadings = ({
  collected,
  layout,
  previous,
}: ResolveOutlineHeadingsOptions): HeadingInfo[] => {
  const previousById = new Map(
    previous.map((heading) => [heading.id, heading]),
  );
  let changed = collected.length !== previous.length;
  const next: HeadingInfo[] = [];
  for (const [index, heading] of collected.entries()) {
    const before = previousById.get(heading.id);
    const pageNumber =
      (layout && layoutPageNumber(layout, heading.pmPos)) ??
      before?.pageNumber ??
      null;
    const resolved: HeadingInfo = {
      id: heading.id,
      text: heading.text,
      level: heading.level,
      pmPos: heading.pmPos,
      pageNumber,
    };
    const kept = before && sameHeading(before, resolved) ? before : resolved;
    if (kept !== previous[index]) {
      changed = true;
    }
    next.push(kept);
  }
  return changed ? next : previous;
};
