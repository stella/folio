/**
 * Floating images lifted out of paragraphs and placed on the page.
 *
 * A floating run stays in its paragraph's run list but paints in a page-level
 * layer, which is why `renderLine` skips it (`renderParagraph.ts:2716`) and
 * `renderPage` re-extracts it. The same split happens here, with one deliberate
 * fix: `renderPage` extracts per fragment, so a paragraph split into two
 * fragments on one page paints its floats twice. Extraction here is keyed on
 * the run, so it cannot.
 */

import type { ImageRun, Page, ParagraphBlock } from "../../layout-engine/types";
import { isFloatingImageRun } from "../../layout-engine/types";
import {
  resolveAnchoredImagePosition,
  type PageGeometry,
} from "../../layout-painter/anchoredImagePosition";
import type { DisplayPrimitive } from "../types";
import type { BuildContext } from "./buildContext";
import { paintImage } from "./imagePrimitives";

export const pageGeometryOf = (page: Page): PageGeometry => ({
  pageWidth: page.size.w,
  pageHeight: page.size.h,
  marginLeft: page.margins.left,
  marginTop: page.margins.top,
  marginRight: page.margins.right,
  marginBottom: page.margins.bottom,
  ...(page.authoredMargins === undefined ? {} : { authoredMargins: page.authoredMargins }),
  contentWidth: page.size.w - page.margins.left - page.margins.right,
  contentHeight: page.size.h - page.margins.top - page.margins.bottom,
});

export type FloatingImagePlacement = {
  readonly run: ImageRun;
  /** `w:wrap behind` paints under the body text; everything else paints over it. */
  readonly behindDoc: boolean;
  readonly primitives: readonly DisplayPrimitive[];
};

export type CollectFloatingImagesOptions = {
  readonly block: ParagraphBlock;
  /** The paragraph fragment's y, page-absolute. */
  readonly fragmentYPx: number;
  readonly geometry: PageGeometry;
  readonly context: BuildContext;
  /** Runs already placed on this page, so a split paragraph cannot double-paint. */
  readonly placed: Set<ImageRun>;
};

export const collectFloatingImages = ({
  block,
  fragmentYPx,
  geometry,
  context,
  placed,
}: CollectFloatingImagesOptions): readonly FloatingImagePlacement[] => {
  const placements: FloatingImagePlacement[] = [];

  for (const run of block.runs) {
    if (run.kind !== "image" || !isFloatingImageRun(run) || placed.has(run)) {
      continue;
    }
    placed.add(run);

    // `resolveAnchoredImagePosition` answers in content-area coordinates.
    const anchor = resolveAnchoredImagePosition(run, fragmentYPx - geometry.marginTop, geometry);
    placements.push({
      run,
      behindDoc: run.wrapType === "behind",
      primitives: paintImage({
        source: run,
        rect: {
          xPx: geometry.marginLeft + anchor.x,
          yPx: geometry.marginTop + anchor.y,
          widthPx: run.width,
          heightPx: run.height,
        },
        context,
        label: "floating image",
      }),
    });
  }

  return placements;
};
