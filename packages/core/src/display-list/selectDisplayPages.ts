/**
 * A display list narrowed to some of its pages.
 *
 * Page indices appear in three places besides `pages`: internal link targets,
 * outline entries, and `unsupported` reports. Each is re-indexed onto the kept
 * pages; a link or outline entry that points at a dropped page is removed, so
 * a backend never meets an index into a page it was not given.
 */

import type { DisplayLink, DisplayList, DisplayPage } from "./types";

/** Keep `pageIndices` (0-based, in the order given; duplicates are dropped). */
export const selectDisplayPages = (
  list: DisplayList,
  pageIndices: readonly number[],
): DisplayList => {
  const kept = [...new Set(pageIndices)].filter(
    (index) => Number.isInteger(index) && index >= 0 && index < list.pages.length,
  );
  const remap = new Map(kept.map((original, next) => [original, next]));

  const relink = (link: DisplayLink): DisplayLink[] => {
    if (link.target.kind !== "page") return [link];
    const pageIndex = remap.get(link.target.pageIndex);
    return pageIndex === undefined ? [] : [{ ...link, target: { ...link.target, pageIndex } }];
  };

  const pages = kept.flatMap((index): DisplayPage[] => {
    const page = list.pages[index];
    return page === undefined ? [] : [{ ...page, links: page.links.flatMap(relink) }];
  });

  return {
    ...list,
    pages,
    outline: list.outline.flatMap((entry) => {
      const pageIndex = remap.get(entry.pageIndex);
      return pageIndex === undefined ? [] : [{ ...entry, pageIndex }];
    }),
    unsupported: list.unsupported.flatMap((entry) => {
      const pageIndex = remap.get(entry.pageIndex);
      return pageIndex === undefined ? [] : [{ ...entry, pageIndex }];
    }),
  };
};
