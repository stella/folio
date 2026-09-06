/**
 * Where a header or footer story sits on the page.
 *
 * The painter wraps each story in a box it also uses as a hover and
 * double-click target, and that box's height is deliberately not the story's
 * visual extent (`renderPage.ts`, `interactiveHeaderHeight`). None of that is
 * paint, so none of it is here: what remains is the origin the story's first
 * in-flow line paints at, which is `headerDistance` below the page top and
 * `footerDistance` plus the story's own flow height above the page bottom.
 *
 * The painter clips a story only when it already fits its box, so the clip
 * never removes a glyph and no `clipGroup` is emitted for it.
 */

import type { HeaderFooterContent, Page } from "../../layout-engine/types";
import type { DisplayPrimitive } from "../types";
import type { BuildContext } from "./buildContext";
import { paintHeaderFooterBlocks } from "./storyPrimitives";

/** `renderPage.ts`: the fallback when neither the caller nor `w:pgMar` states one. */
const DEFAULT_DISTANCE_PX = 48;

export type HeaderFooterSection = "header" | "footer";

export type HeaderFooterPaintOptions = {
  readonly page: Page;
  readonly section: HeaderFooterSection;
  readonly content: HeaderFooterContent;
  /** Overrides `w:pgMar`'s `w:header` / `w:footer` for this page. */
  readonly distancePx?: number;
  readonly context: BuildContext;
};

export const paintHeaderFooter = ({
  page,
  section,
  content,
  distancePx,
  context,
}: HeaderFooterPaintOptions): readonly DisplayPrimitive[] => {
  if (content.blocks.length === 0) {
    return [];
  }

  const distance =
    distancePx ??
    (section === "header" ? page.margins.header : page.margins.footer) ??
    DEFAULT_DISTANCE_PX;
  const originYPx = section === "header" ? distance : page.size.h - distance - content.height;

  return paintHeaderFooterBlocks({
    blocks: content.blocks,
    measures: content.measures,
    xPx: page.margins.left,
    yPx: originYPx,
    widthPx: page.size.w - page.margins.left - page.margins.right,
    context: { ...context, story: section },
    label: content.rId === undefined ? section : `${section} ${content.rId}`,
  });
};
