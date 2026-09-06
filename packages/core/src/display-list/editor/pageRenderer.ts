/**
 * Which renderer paints the editor's pages.
 *
 * The names are data rather than a boolean because this is a "which kind"
 * question that has already had two answers and may have a third; a
 * `useDisplayList` flag would have to be migrated the moment it does.
 */

import type { RenderPageOptions } from "../../layout-painter/renderPage";
import type { PageFurnitureInputs } from "../build/furniture";

export const PAGE_RENDERER = {
  /** The painter that has always painted the editor. */
  legacy: "legacy",
  /** Builds the paint IR from the layout and paints the pages from it. */
  displayList: "display-list",
} as const;

export type PageRendererName = (typeof PAGE_RENDERER)[keyof typeof PAGE_RENDERER];

export const isPageRendererName = (value: unknown): value is PageRendererName =>
  value === PAGE_RENDERER.legacy || value === PAGE_RENDERER.displayList;

/**
 * Read the display-list producer's furniture inputs out of the render options
 * the pipeline already computed.
 *
 * Both sides describe the same constructs, and the pipeline is the only place
 * that knows how to derive them from a package, so the display-list path reads
 * what the existing path receives rather than deriving them a second time. A
 * second derivation is how the two renderers would come to disagree about
 * which header a page carries.
 */
export const displayListFurnitureFrom = (options: RenderPageOptions): PageFurnitureInputs => ({
  ...(options.pageBorders === undefined ? {} : { pageBorders: options.pageBorders }),
  ...(options.theme === undefined ? {} : { theme: options.theme }),
  ...(options.watermark === undefined ? {} : { watermark: options.watermark }),
  ...(options.watermarkByHeaderRId === undefined
    ? {}
    : { watermarkByHeaderRId: options.watermarkByHeaderRId }),
  ...(options.watermarkImageSrc === undefined
    ? {}
    : { watermarkImageSrc: options.watermarkImageSrc }),
  ...(options.headerContent === undefined ? {} : { headerContent: options.headerContent }),
  ...(options.footerContent === undefined ? {} : { footerContent: options.footerContent }),
  ...(options.headerContentByRId === undefined
    ? {}
    : { headerContentByRId: options.headerContentByRId }),
  ...(options.footerContentByRId === undefined
    ? {}
    : { footerContentByRId: options.footerContentByRId }),
  ...(options.firstPageHeaderContent === undefined
    ? {}
    : { firstPageHeaderContent: options.firstPageHeaderContent }),
  ...(options.firstPageFooterContent === undefined
    ? {}
    : { firstPageFooterContent: options.firstPageFooterContent }),
  ...(options.titlePg === undefined ? {} : { titlePg: options.titlePg }),
  ...(options.headerDistance === undefined ? {} : { headerDistancePx: options.headerDistance }),
  ...(options.footerDistance === undefined ? {} : { footerDistancePx: options.footerDistance }),
});
