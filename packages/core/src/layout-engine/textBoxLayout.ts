/**
 * Places text boxes onto pages: inline and anchored placement, page-pinned
 * wrap bands, and host-paragraph line offsets.
 */

import { emuToPixels } from "../utils/units";
import { createPaginator } from "./paginator";
import { bandFragmentX, bandTopContentY, isPageFrameRelativeAnchor } from "./textBoxFlow";
import { floatingTextBoxReservesBand } from "./types";
import type { TextBoxBlock, TextBoxMeasure, TextBoxFragment } from "./types";

type LayoutTextBoxOptions = {
  paginator: ReturnType<typeof createPaginator>;
  sectionMarginTop: number;
  sectionPageHeight: number;
  sectionMarginBottom: number;
};

const TEXT_BOX_ANCHOR_BLOCK_ID = Symbol.for("stll.textBoxAnchorBlockId");

const readTextBoxAnchorBlockId = (block: TextBoxBlock): unknown =>
  Reflect.get(block, TEXT_BOX_ANCHOR_BLOCK_ID);

/**
 * Layout a text box block onto pages.
 */
export function layoutTextBox(
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  { paginator, sectionMarginTop, sectionPageHeight, sectionMarginBottom }: LayoutTextBoxOptions,
): void {
  // A page/margin-pinned topAndBottom band (e.g. a title banner) floats to the
  // top of its page; the reserved band in the measure pass pushes body text
  // below it (see extractFloatingZones in PagedEditor). It must not also consume
  // flow at its anchor, or the box height would be reserved twice. Place it at
  // the page content top without advancing the cursor. eigenpal #694.
  if (isPagePinnedBandTextBox(block)) {
    const state = paginator.getCurrentState();
    // Position the box at the same content-Y the measure pass reserved its band
    // at. The measure pass uses the section top margin (not a page's
    // first-page margin), so use the same value here; `bandTopContentY` is
    // content-relative, and `fragment.y` is page-absolute, so add the page's
    // own top margin (`state.topMargin`) to convert. Using `state.topMargin`
    // inside the resolver instead would desync the box from its band on a
    // title page whose first-page top margin differs from the section margin.
    const bandTop = bandTopContentY(block.position?.vertical, {
      pageHeight: sectionPageHeight,
      marginTop: sectionMarginTop,
      marginBottom: sectionMarginBottom,
      boxHeight: measure.height,
    });
    // Honor the box's horizontal anchor (align center/right, page-relative
    // offset) instead of always pinning to the column's left edge. The band is
    // full-width regardless, so this only moves where the box paints.
    const x = bandFragmentX(block.position?.horizontal, {
      pageWidth: state.page.size.w,
      marginLeft: state.page.margins.left,
      marginRight: state.page.margins.right,
      activeColumnLeft: paginator.getColumnX(state.columnIndex),
      activeColumnWidth: paginator.columnWidth,
      boxWidth: measure.width,
    });
    const fragment: TextBoxFragment = {
      kind: "textBox",
      blockId: block.id,
      x,
      y:
        block.position?.vertical?.relativeTo === "page"
          ? sectionMarginTop + bandTop
          : state.topMargin + bandTop,
      width: measure.width,
      height: measure.height,
      isPositioned: true,
      ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
      ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
    };
    paginator.addUnflowedFragment(fragment);
    return;
  }

  // Any explicitly positioned textbox is an anchored object, including
  // paragraph/line-relative anchors. Place it from the current text anchor
  // without advancing normal flow; otherwise several shapes owned by one
  // shape-only paragraph stack vertically and push the body onto another page.
  if (block.position !== undefined) {
    const state = paginator.getCurrentState();
    const x = bandFragmentX(block.position.horizontal, {
      pageWidth: state.page.size.w,
      marginLeft: state.page.margins.left,
      marginRight: state.page.margins.right,
      activeColumnLeft: paginator.getColumnX(state.columnIndex),
      activeColumnWidth: paginator.columnWidth,
      boxWidth: measure.width,
    });
    const vertical = block.position.vertical;
    const anchorBlockId = readTextBoxAnchorBlockId(block);
    const anchorParagraph =
      vertical?.relativeTo === "paragraph" && typeof anchorBlockId === "string"
        ? state.page.fragments.find(
            (fragment) => fragment.kind === "paragraph" && fragment.blockId === anchorBlockId,
          )
        : undefined;
    const bandGeometry = {
      pageHeight: sectionPageHeight,
      marginTop: sectionMarginTop,
      marginBottom: sectionMarginBottom,
      boxHeight: measure.height,
    };
    let y;
    if (vertical?.relativeTo === "page") {
      y = sectionMarginTop + bandTopContentY(vertical, bandGeometry);
    } else if (isPageFrameRelativeAnchor(vertical?.relativeTo)) {
      y =
        state.topMargin +
        bandTopContentY(vertical, {
          pageHeight: sectionPageHeight,
          marginTop: sectionMarginTop,
          marginBottom: sectionMarginBottom,
          boxHeight: measure.height,
        });
    } else {
      y = (anchorParagraph?.y ?? state.cursorY) + emuToPixels(vertical?.posOffset ?? 0);
    }
    const fragment: TextBoxFragment = {
      kind: "textBox",
      blockId: block.id,
      x,
      y,
      width: measure.width,
      height: measure.height,
      isPositioned: true,
      ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
      ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
    };
    paginator.addUnflowedFragment(fragment);
    return;
  }

  // An inline box occupies its effect extent beyond its own on every side.
  const effect = block.effectExtent;
  const occupiedHeight = measure.height + (effect?.top ?? 0) + (effect?.bottom ?? 0);
  const occupiedWidth = measure.width + (effect?.left ?? 0) + (effect?.right ?? 0);
  const host = block.hostParagraph;

  const fragment: TextBoxFragment = {
    kind: "textBox",
    blockId: block.id,
    x: 0,
    y: 0,
    width: measure.width,
    height: measure.height,
    ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
    ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
  };

  // The box is the line of its host paragraph, so the host's spacing
  // surrounds it and collapses with its neighbours' like any paragraph's.
  const result = paginator.addFragment(
    fragment,
    occupiedHeight,
    host?.spacing?.before ?? 0,
    host?.spacing?.after ?? 0,
  );
  fragment.x =
    result.x +
    hostParagraphLineOffset(host, paginator.columnWidth, occupiedWidth) +
    (effect?.left ?? 0);
  fragment.y = result.y + (effect?.top ?? 0);
}

/**
 * Where a line holding only an object of `width` starts within a column:
 * after the host paragraph's start and first-line indents, then aligned in
 * what remains before its end indent.
 */
function hostParagraphLineOffset(
  host: TextBoxBlock["hostParagraph"],
  columnWidth: number,
  width: number,
): number {
  if (!host) {
    return 0;
  }
  const indent = host.indent;
  const start = (indent?.left ?? 0) + (indent?.firstLine ?? 0) - (indent?.hanging ?? 0);
  const slack = Math.max(0, columnWidth - start - (indent?.right ?? 0) - width);
  switch (host.alignment) {
    case "center":
      return start + slack / 2;
    case "right":
      return start + slack;
    case "left":
    case "justify":
    case undefined:
      return start;
    default:
      host.alignment satisfies never;
      return start;
  }
}

/**
 * A topAndBottom text box whose vertical anchor pins it to the page frame
 * (page/margin/margin-strip) — it floats to a fixed page position rather than
 * flowing in document order. Must agree with the measure pass's band extraction
 * (extractFloatingZones), which uses the same predicate. eigenpal #694.
 */
function isPagePinnedBandTextBox(block: TextBoxBlock): boolean {
  if (!floatingTextBoxReservesBand(block)) {
    return false;
  }
  return isPageFrameRelativeAnchor(block.position?.vertical?.relativeTo);
}
