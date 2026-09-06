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
 * ## Why a page is built when it is painted
 *
 * `renderPage` is called per page, so building the whole list per call would be
 * quadratic in page count. Building it once per layout run is not right either:
 * the container paints a window of pages and skips every page whose fingerprint
 * has not changed, so a document of ninety pages would build ninety to paint
 * three, on every keystroke. The builder computes what belongs to the document
 * once and each page when it is asked for, so the work follows the screen.
 */

import { panic } from "better-result";

import { createDisplayListBuilder } from "../build/buildDisplayList";
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
  readonly paintPage: (request: { readonly page: Page }) => HTMLElement | null;
  /**
   * The list as far as the painter has built it, for tests and diagnostics.
   * Pages nobody painted are absent, because nobody built them.
   */
  readonly list: () => DisplayList;
};

/**
 * Page numbers are the engine's own and may restart or skip across sections,
 * so a display page is found by its position in the layout rather than by its
 * number. Both sequences come from the same `Layout`, so they align by
 * construction; a mismatch is a programming error, not a document one.
 */
const indexByPageNumber = (
  layout: BuildDisplayListOptions["layout"],
): ReadonlyMap<number, number> => new Map(layout.pages.map((page, index) => [page.number, index]));

export const createDisplayListPagePainter = (
  options: DisplayListPagePainterOptions,
): DisplayListPagePainter => {
  const { doc, ...build } = options;
  const builder = createDisplayListBuilder(build);
  const indexOf = indexByPageNumber(build.layout);

  const paintPage = ({ page }: { readonly page: Page }): HTMLElement | null => {
    const index = indexOf.get(page.number);
    if (index === undefined) {
      // The layout grew a page after this painter was made. Yielding paints the
      // page through the existing renderer, which is wrong-looking rather
      // than missing, and the next layout run makes a new painter.
      return null;
    }
    const displayPage = builder.pageAt(index);
    if (displayPage === undefined) {
      panic(`display list page ${String(index)} is indexed but absent`);
    }
    // The tables are read after the page is built, never before: building it is
    // what puts its faces and images in them.
    const { fonts, images } = builder.snapshot();
    return renderDisplayPageToDom(displayPage, { doc, fonts, images, pageIndex: index });
  };

  return { paintPage, list: builder.snapshot };
};
