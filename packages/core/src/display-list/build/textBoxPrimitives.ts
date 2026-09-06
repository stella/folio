/**
 * Text-box fragments → primitives: fill, outline, then the inner block story.
 *
 * The box's own content flow is `layoutTextBoxContent`, the same helper the
 * measurer used, so the paragraphs land where their heights were reserved.
 */

import { layoutTextBoxContent } from "../../layout-engine/measure/textBoxParagraphLayout";
import { DEFAULT_TEXTBOX_MARGINS } from "../../layout-engine/types";
import type { TextBoxBlock, TextBoxFragment, TextBoxMeasure } from "../../layout-engine/types";
import type { DisplayPrimitive } from "../types";
import { HIT_REGION_KINDS } from "../primitives";
import type { BuildContext } from "./buildContext";
import { blockRegion, type PageComposer } from "./regions";
import { parseDisplayColor } from "./colors";
import { paintParagraphFragment } from "./paragraphPrimitives";
import { paintTableBlock } from "./tablePrimitives";
import { resolveBorderStroke } from "./strokes";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

export type TextBoxPaintOptions = {
  readonly composer: PageComposer;
  readonly fragment: TextBoxFragment;
  readonly block: TextBoxBlock;
  readonly measure: TextBoxMeasure;
  readonly context: BuildContext;
};

export const paintTextBoxFragment = ({
  composer,
  fragment,
  block,
  measure,
  context,
}: TextBoxPaintOptions): void => {
  const primitives: DisplayPrimitive[] = [];
  const box = {
    xPx: fragment.x,
    yPx: fragment.y,
    widthPx: fragment.width,
    heightPx: fragment.height,
  };

  if (block.fillColor) {
    const fill = parseDisplayColor(block.fillColor);
    if (fill) {
      primitives.push({ kind: "rect", rect: box, fill });
    } else {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.unresolvedColor,
        context.pageIndex,
        `text box fill ${block.fillColor}`,
      );
    }
  }

  let outlineWidthPx = 0;
  if (block.outlineWidth !== undefined && block.outlineWidth > 0) {
    const { stroke, unresolvedColor } = resolveBorderStroke({
      width: block.outlineWidth,
      style: block.outlineStyle ?? "solid",
      color: block.outlineColor ?? "#000000",
    });
    if (unresolvedColor !== undefined) {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.unresolvedColor,
        context.pageIndex,
        `text box outline ${unresolvedColor}`,
      );
    }
    if (stroke) {
      outlineWidthPx = stroke.thicknessPx;
      // `box-sizing: border-box`: the outline paints inside the authored box.
      const inset = stroke.thicknessPx / 2;
      primitives.push({
        kind: "rect",
        rect: {
          xPx: box.xPx + inset,
          yPx: box.yPx + inset,
          widthPx: Math.max(0, box.widthPx - stroke.thicknessPx),
          heightPx: Math.max(0, box.heightPx - stroke.thicknessPx),
        },
        stroke,
      });
    }
  }

  const margins = block.margins ?? DEFAULT_TEXTBOX_MARGINS;
  // The measurer wrapped against `width - margins`, ignoring the outline, so
  // paint must use the same width or the text would rewrap.
  const innerWidthPx = fragment.width - margins.left - margins.right;
  const layout = layoutTextBoxContent(block.content, measure.innerMeasures);

  let cursorYPx = box.yPx + outlineWidthPx + margins.top;
  for (let index = 0; index < block.content.length; index += 1) {
    const contentBlock = block.content[index];
    const contentMeasure = measure.innerMeasures[index];
    const placement = layout.placements[index];
    if (!contentBlock || !contentMeasure || !placement) {
      continue;
    }
    cursorYPx += placement.leadingSpacing;

    if (contentBlock.kind === "paragraph" && contentMeasure.kind === "paragraph") {
      const previous = block.content[index - 1];
      const next = block.content[index + 1];
      composer.push(primitives);
      primitives.length = 0;
      const contentFragment = {
        kind: "paragraph",
        blockId: contentBlock.id,
        x: box.xPx + outlineWidthPx + margins.left,
        y: cursorYPx,
        width: innerWidthPx,
        height: placement.contentHeight,
        fromLine: 0,
        toLine: contentMeasure.lines.length,
      } as const;
      composer.region(
        blockRegion({ fragment: contentFragment, kind: HIT_REGION_KINDS.paragraph, context }),
        () => {
          paintParagraphFragment({
            composer,
            fragment: contentFragment,
            block: contentBlock,
            measure: contentMeasure,
            context,
            ...(previous?.kind === "paragraph" && previous.attrs?.borders !== undefined
              ? { prevBorders: previous.attrs.borders }
              : {}),
            ...(next?.kind === "paragraph" && next.attrs?.borders !== undefined
              ? { nextBorders: next.attrs.borders }
              : {}),
          });
        },
      );
    } else if (contentBlock.kind === "table" && contentMeasure.kind === "table") {
      composer.push(primitives);
      primitives.length = 0;
      const contentFragment = {
        blockId: contentBlock.id,
        x: box.xPx + outlineWidthPx + margins.left,
        y: cursorYPx,
        width: contentMeasure.totalWidth,
        height: contentMeasure.totalHeight,
      };
      composer.region(
        blockRegion({ fragment: contentFragment, kind: HIT_REGION_KINDS.table, context }),
        () => {
          paintTableBlock({
            composer,
            block: contentBlock,
            measure: contentMeasure,
            xPx: contentFragment.x,
            yPx: contentFragment.y,
            context,
            paintTextBox: paintTextBoxFragment,
          });
        },
      );
    } else {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.fragmentKind,
        context.pageIndex,
        `text box content of kind ${contentBlock.kind} is not painted`,
      );
    }

    cursorYPx += placement.contentHeight;
  }

  if (block.verticalAlign === "middle" || block.verticalAlign === "bottom") {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.textBoxVerticalAlign,
      context.pageIndex,
      `text box anchor ${block.verticalAlign} paints from the top`,
    );
  }

  composer.push(primitives);
};
