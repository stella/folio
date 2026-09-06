/**
 * Paint the editor's pages from the display list instead of from `Layout`.
 *
 * This is a composition root, not a backend: it is the one place that knows
 * both the producer and the DOM backend, which is why it may import each of
 * them and why neither may import it.
 *
 * ## What it replaces, and what it deliberately does not
 *
 * `renderPages` owns two separable jobs. One is *painting* a page, which is
 * what diverges between two renderers and what the display list exists to make
 * single. The other is the page *container*: shells, virtualization, the
 * fingerprint comparison that skips an unchanged page, the intersection
 * observer, the painted event. That second job is renderer-independent, so
 * this module supplies only a `paintPage` and leaves the container alone.
 * Reimplementing the container behind a second renderer would be exactly the
 * second copy the display list was introduced to remove, and it would silently
 * lose incremental repaint on documents past the virtualization threshold.
 *
 * ## Why the list is built once
 *
 * `renderPage` is called per page, and building the whole display list per
 * call would be quadratic in page count. The list is built once from the
 * layout the pipeline just produced, and `paintPage` looks its page up.
 */

import { panic } from "better-result";

import { buildDisplayList } from "../build/buildDisplayList";
import type { BuildDisplayListOptions } from "../build/buildDisplayList";
import { renderDisplayPageToDom } from "../dom/renderDisplayListToDom";
import type { DisplayList } from "../types";
import type { Page } from "../../layout-engine/types";

export type DisplayListPagePainterOptions = BuildDisplayListOptions & {
  /** Document to create elements in. */
  readonly doc: Document;
};

export type DisplayListPagePainter = {
  /** Matches `RenderPageOptions.paintPage`. */
  readonly paintPage: (page: Page) => HTMLElement | null;
  /** The list every page was painted from, for tests and diagnostics. */
  readonly list: DisplayList;
};

/**
 * Page numbers are the engine's own and may restart or skip across sections,
 * so a display page is found by its position in the layout rather than by its
 * number. Both sequences come from the same `Layout`, so they align by
 * construction; a mismatch is a programming error, not a document one.
 */
const indexByPageNumber = (list: DisplayList): ReadonlyMap<number, number> =>
  new Map(list.pages.map((page, index) => [page.pageNumber, index]));

export const createDisplayListPagePainter = (
  options: DisplayListPagePainterOptions,
): DisplayListPagePainter => {
  const { doc, ...build } = options;
  const list = buildDisplayList(build);
  const indexOf = indexByPageNumber(list);

  const paintPage = (page: Page): HTMLElement | null => {
    const index = indexOf.get(page.number);
    if (index === undefined) {
      // The layout grew a page after the list was built. Yielding paints the
      // page through the existing renderer, which is wrong-looking rather
      // than missing, and the next layout run rebuilds the list.
      return null;
    }
    const displayPage = list.pages.at(index);
    if (displayPage === undefined) {
      panic(`display list page ${String(index)} is indexed but absent`);
    }
    return renderDisplayPageToDom(displayPage, {
      doc,
      fonts: list.fonts,
      images: list.images,
      pageIndex: index,
    });
  };

  return { paintPage, list };
};
