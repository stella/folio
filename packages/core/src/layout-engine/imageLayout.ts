/**
 * Places block-level and anchored images onto pages.
 */

import { createPaginator } from "./paginator";
import type { ImageBlock, ImageMeasure, ImageFragment } from "./types";

/**
 * Layout an image block onto pages.
 */
export function layoutImage(
  block: ImageBlock,
  measure: ImageMeasure,
  paginator: ReturnType<typeof createPaginator>,
): void {
  // Handle anchored images differently
  if (block.anchor?.isAnchored) {
    layoutAnchoredImage(block, measure, paginator);
    return;
  }

  // Inline image - ensure it fits (plus any leading skip past a page band)
  const bandSkip = measure.bandSkipBefore ?? 0;
  const state = paginator.ensureFits(bandSkip + measure.height);

  const fragment: ImageFragment = {
    kind: "image",
    blockId: block.id,
    x: paginator.getColumnX(state.columnIndex),
    y: 0, // Will be set by addFragment
    width: measure.width,
    height: measure.height,
    ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
    ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
  };

  const result = paginator.addFragment(fragment, measure.height, bandSkip, 0);
  fragment.y = result.y;
}

/**
 * Layout an anchored (floating) image.
 */
function layoutAnchoredImage(
  block: ImageBlock,
  measure: ImageMeasure,
  paginator: ReturnType<typeof createPaginator>,
): void {
  const state = paginator.getCurrentState();
  const anchor = block.anchor;
  if (!anchor) {
    return;
  }

  // Position based on anchor offsets
  const x = anchor.offsetH ?? paginator.getColumnX(state.columnIndex);
  const y = anchor.offsetV ?? state.cursorY;

  const fragment: ImageFragment = {
    kind: "image",
    blockId: block.id,
    x,
    y,
    width: measure.width,
    height: measure.height,
    ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
    ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
    isAnchored: true,
    zIndex: anchor.behindDoc ? -1 : 1,
  };

  paginator.addUnflowedFragment(fragment);
}
