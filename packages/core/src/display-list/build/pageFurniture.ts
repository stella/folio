/**
 * Page furniture: everything on a page that is not a body fragment.
 *
 * `renderPage.ts` is the authority on the order and on the geometry, and it
 * takes most of this from `RenderPageOptions` rather than from the `Layout`:
 * page borders, watermarks, footnote bodies and header/footer stories all reach
 * the DOM painter as render options the layout never carried. The builder is
 * handed the same inputs, so a construct the caller supplies is painted and a
 * construct the layout shows evidence of but the caller withheld is reported.
 * The furniture that is derivable from the `Layout` alone — the page
 * background, the column separators and the footnote separator rule — is always
 * painted.
 */

import { FOOTNOTE_SEPARATOR_HEIGHT } from "../../layout-engine/types";
import type { FootnoteContent, Page } from "../../layout-engine/types";
import {
  calculateFootnoteAreaRenderHeight,
  type FootnoteRenderItem,
} from "../../layout-painter/renderPage";
import type { DisplayColor, DisplayPrimitive } from "../types";
import type { BuildContext } from "./buildContext";
import { DOC_CANVAS_TEXT } from "./colors";
import { paintFootnoteBlocks } from "./storyPrimitives";
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

export type FootnoteAreaPaintOptions = {
  readonly page: Page;
  readonly context: BuildContext;
  /**
   * The bodies, by `w:footnote` id. Absent, or missing an id the page carries,
   * means the caller did not supply that body: the band still opens with its
   * rule and reserves its height, and the missing bodies are named.
   */
  readonly contentById?: ReadonlyMap<number, FootnoteContent>;
};

/**
 * The footnote band at the foot of the body column: the separator rule, then
 * one body per note the paginator put on this page.
 *
 * The band's top comes from the reservation the paginator made, clamped up to
 * what the bodies actually need, so a stack that under-reserved by a pixel ends
 * at the page bottom instead of spilling past it.
 */
export const paintFootnoteArea = ({
  page,
  context,
  contentById,
}: FootnoteAreaPaintOptions): readonly DisplayPrimitive[] => {
  const reservedHeightPx = page.footnoteReservedHeight ?? 0;
  const noteIds = page.footnoteIds ?? [];
  if (reservedHeightPx <= 0 || noteIds.length === 0) {
    return [];
  }

  const bodies = noteIds.map((noteId) => ({ noteId, content: contentById?.get(noteId) }));
  const missing = bodies.flatMap(({ noteId, content }) => (content === undefined ? [noteId] : []));
  if (missing.length > 0) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.footnoteContent,
      context.pageIndex,
      `footnote bodies ${missing.join(", ")} were not supplied to the builder, so the band reserves their height and paints nothing`,
    );
  }

  const items: FootnoteRenderItem[] = bodies.map(({ noteId, content }) => ({
    noteId,
    displayNumber: String(content?.displayNumber ?? noteId),
    ...(content === undefined
      ? {}
      : {
          content: { blocks: content.blocks, measures: content.measures, height: content.height },
        }),
  }));

  const contentHeightPx = page.size.h - page.margins.top - page.margins.bottom;
  const contentWidthPx = page.size.w - page.margins.left - page.margins.right;
  const bandHeightPx = Math.max(reservedHeightPx, calculateFootnoteAreaRenderHeight(items));
  const areaTopPx = page.margins.top + Math.max(-page.margins.top, contentHeightPx - bandHeightPx);

  const primitives: DisplayPrimitive[] = [
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

  for (const [index, { content }] of bodies.entries()) {
    if (content === undefined) {
      continue;
    }
    // The band's own height function over the notes above this one: the
    // separator slot plus each preceding entry and its margin, which is exactly
    // this entry's offset from the band top. Taking it from the same helper the
    // clamp above uses keeps a note that was not supplied from shifting the
    // ones below it.
    const offsetPx = calculateFootnoteAreaRenderHeight(items.slice(0, index));
    primitives.push(
      ...paintFootnoteBlocks({
        blocks: content.blocks,
        measures: content.measures,
        xPx: page.margins.left,
        yPx: areaTopPx + offsetPx,
        widthPx: contentWidthPx,
        context: { ...context, story: "footnote" },
        label: `footnote ${content.displayNumber}`,
      }),
    );
  }

  return primitives;
};
