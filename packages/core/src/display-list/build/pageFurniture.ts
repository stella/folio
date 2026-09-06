/**
 * Page furniture: everything on a page that is not a body fragment.
 *
 * `renderPage.ts` is the authority on the order, and it takes most of this from
 * `RenderPageOptions` rather than from the `Layout`: page borders, watermarks,
 * footnote bodies and header/footer stories all reach the DOM painter as render
 * options that the layout never carried. `buildDisplayList` is handed a
 * `Layout` and a `BlockLookup`, so those four are not reachable from here and
 * are reported through `unsupported` instead of being silently absent. The
 * furniture that *is* derivable — the page background, the column separators
 * and the footnote separator rule — is painted.
 */

import { FOOTNOTE_SEPARATOR_HEIGHT } from "../../layout-engine/types";
import type { Page } from "../../layout-engine/types";
import type { DisplayColor, DisplayPrimitive } from "../types";
import type { BuildContext } from "./buildContext";
import { DOC_CANVAS_TEXT } from "./colors";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

/** `renderPage.ts:2049`: a hairline in the canvas ink colour. */
const COLUMN_SEPARATOR_WIDTH_PX = 0.5;

/** `renderPage.ts:1354-1360`: a 0.5px rule across a third of the column, centred in its 12px slot. */
const FOOTNOTE_RULE_THICKNESS_PX = 0.5;
const FOOTNOTE_RULE_WIDTH_FRACTION = 0.33;
const FOOTNOTE_RULE_MARGIN_PX = (FOOTNOTE_SEPARATOR_HEIGHT - FOOTNOTE_RULE_THICKNESS_PX) / 2;

export const paintPageBackground = (page: Page, fill: DisplayColor): DisplayPrimitive => ({
  kind: "rect",
  rect: { xPx: 0, yPx: 0, widthPx: page.size.w, heightPx: page.size.h },
  fill,
});

/**
 * Vertical rules between newspaper columns. The painter assumes equal columns
 * with a uniform gap (it ignores `columns.widths` / `columns.gaps`), and so
 * does this: diverging here would put the export's rules somewhere the editor
 * never drew them.
 */
export const paintColumnSeparators = (page: Page): readonly DisplayPrimitive[] => {
  const columns = page.columns;
  if (!columns?.separator || columns.count <= 1) {
    return [];
  }

  const contentWidthPx = page.size.w - page.margins.left - page.margins.right;
  const contentHeightPx = page.size.h - page.margins.top - page.margins.bottom;
  const columnWidthPx = (contentWidthPx - (columns.count - 1) * columns.gap) / columns.count;

  const primitives: DisplayPrimitive[] = [];
  for (let column = 0; column < columns.count - 1; column += 1) {
    const offsetPx = (column + 1) * columnWidthPx + column * columns.gap + columns.gap / 2;
    primitives.push({
      kind: "rect",
      rect: {
        xPx: page.margins.left + offsetPx,
        yPx: page.margins.top,
        widthPx: COLUMN_SEPARATOR_WIDTH_PX,
        heightPx: contentHeightPx,
      },
      fill: DOC_CANVAS_TEXT,
    });
  }
  return primitives;
};

/**
 * The footnote separator rule.
 *
 * The bodies are not painted: `FootnoteContent` reaches the DOM painter through
 * `RenderPageOptions.footnoteArea`, and nothing in the `Layout` carries it. The
 * reservation itself is a layout fact, so the rule that opens the reserved band
 * is drawn and the missing bodies are named.
 */
export const paintFootnoteArea = (
  page: Page,
  context: BuildContext,
): readonly DisplayPrimitive[] => {
  const reservedHeightPx = page.footnoteReservedHeight ?? 0;
  const noteCount = page.footnoteIds?.length ?? 0;
  if (reservedHeightPx <= 0 || noteCount === 0) {
    return [];
  }

  context.unsupported.report(
    UNSUPPORTED_CONSTRUCT.footnoteContent,
    context.pageIndex,
    `${noteCount} footnote bodies are not painted: FootnoteContent reaches the painter through render options, not through Layout`,
  );

  const contentHeightPx = page.size.h - page.margins.top - page.margins.bottom;
  const contentWidthPx = page.size.w - page.margins.left - page.margins.right;
  const areaTopPx =
    page.margins.top + Math.max(-page.margins.top, contentHeightPx - reservedHeightPx);

  return [
    {
      kind: "rect",
      rect: {
        xPx: page.margins.left,
        yPx: areaTopPx + FOOTNOTE_RULE_MARGIN_PX,
        widthPx: contentWidthPx * FOOTNOTE_RULE_WIDTH_FRACTION,
        heightPx: FOOTNOTE_RULE_THICKNESS_PX,
      },
      fill: DOC_CANVAS_TEXT,
    },
  ];
};

/**
 * Name the furniture the builder's inputs cannot reach, once per page that
 * shows evidence of it. A backend cannot tell a page with no header from a page
 * whose header the producer never received.
 */
export const reportUnreachableFurniture = (page: Page, context: BuildContext): void => {
  if (page.headerFooterRefs !== undefined) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.headerFooterContent,
      context.pageIndex,
      "header and footer stories reach the painter through render options (headerContentByRId), not through Layout",
    );
  }
};
