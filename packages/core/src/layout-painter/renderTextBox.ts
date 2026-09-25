/**
 * Text Box Renderer
 *
 * Renders text box fragments to DOM. Handles:
 * - Background fill color or linear gradient
 * - Border/outline
 * - Internal padding (margins)
 * - Block content inside the box (using pre-measured data)
 */

import { panic } from "better-result";

import { DEFAULT_TEXTBOX_MARGINS } from "../layout-engine/types";
import type {
  TableBlock,
  TableMeasure,
  TextBoxFragment,
  TextBoxBlock,
  TextBoxGradientFill,
  TextBoxMeasure,
} from "../layout-engine/types";
import { presetDashForOutlineAttr } from "../types/documentEnumValues";
import { cssBorderStyleForDash } from "../utils/borderCss";
import { setAuthoredBackgroundColor } from "./documentColors";
import { layoutTextBoxContent } from "../layout-engine/measure/textBoxParagraphLayout";
import { renderParagraphFragment } from "./renderParagraph";
import type { RenderContext } from "./renderUtils";

/**
 * CSS class names for text box elements
 */
export const TEXTBOX_CLASS_NAMES = {
  textBox: "layout-textbox",
};

/**
 * Options for rendering a text box fragment
 */
export type RenderTextBoxFragmentOptions = {
  document?: Document;
  renderTable?: (
    block: TableBlock,
    measure: TableMeasure,
    context: RenderContext,
    document: Document,
    frameWidth?: number,
  ) => HTMLElement;
};

/**
 * The CSS angle of a DrawingML linear gradient over a `width` x `height` box.
 *
 * DrawingML measures `a:lin@ang` clockwise from the positive x axis; CSS
 * measures clockwise from "to top", a quarter turn earlier. A `scaled` angle is
 * stated in the unit square: stretching that square to the box keeps each
 * isoline through the same corners, so the gradient runs perpendicular to the
 * stretched isoline rather than along the stated angle.
 */
export function cssLinearGradientAngle(
  fill: Pick<TextBoxGradientFill, "angle" | "scaled">,
  width: number,
  height: number,
): number {
  const radians = (fill.angle * Math.PI) / 180;
  const along =
    fill.scaled && width > 0 && height > 0
      ? Math.atan2(width * Math.sin(radians), height * Math.cos(radians))
      : radians;
  const degrees = 90 + (along * 180) / Math.PI;
  return ((degrees % 360) + 360) % 360;
}

const formatCssNumber = (value: number): string => String(Math.round(value * 1000) / 1000);

function paintGradientFill(
  style: CSSStyleDeclaration,
  fill: TextBoxGradientFill,
  width: number,
  height: number,
): void {
  const [first] = fill.stops;
  if (first === undefined) {
    return;
  }
  if (fill.stops.length === 1) {
    setAuthoredBackgroundColor(style, first.color);
    return;
  }
  const angle = formatCssNumber(cssLinearGradientAngle(fill, width, height));
  const stops = fill.stops
    .map((stop) => `${stop.color} ${formatCssNumber(stop.offset * 100)}%`)
    .join(", ");
  style.backgroundImage = `linear-gradient(${angle}deg, ${stops})`;
}

/**
 * Render a text box fragment to DOM
 */
export function renderTextBoxFragment(
  fragment: TextBoxFragment,
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  context: RenderContext,
  options: RenderTextBoxFragmentOptions = {},
): HTMLElement {
  const doc = options.document ?? document;

  const containerEl = doc.createElement("div");
  containerEl.className = TEXTBOX_CLASS_NAMES.textBox;

  // Basic styling
  containerEl.style.position = "absolute";
  containerEl.style.width = `${fragment.width}px`;
  containerEl.style.height = `${fragment.height}px`;
  containerEl.style.overflow = block.textWrap === "none" ? "visible" : "hidden";
  containerEl.style.boxSizing = "border-box";

  // Fill color
  if (block.fillColor) {
    setAuthoredBackgroundColor(containerEl.style, block.fillColor);
  } else if (block.fillGradient) {
    paintGradientFill(containerEl.style, block.fillGradient, fragment.width, fragment.height);
  }

  // Border/outline. The node's `outlineStyle` is a DrawingML dash, not a CSS
  // keyword: `sysDash` in the shorthand invalidates the whole declaration and
  // the outline disappears, so it is translated rather than interpolated.
  const outlineDash = presetDashForOutlineAttr(block.outlineStyle) ?? "solid";
  if (block.outlineWidth && block.outlineWidth > 0 && outlineDash !== "none") {
    const color = block.outlineColor || "#000000";
    containerEl.style.border = `${block.outlineWidth}px ${cssBorderStyleForDash(outlineDash)} ${color}`;
  }

  // Internal padding
  const margins = block.margins ?? DEFAULT_TEXTBOX_MARGINS;
  containerEl.style.padding = `${margins.top}px ${margins.right}px ${margins.bottom}px ${margins.left}px`;
  // Distributed and justified anchors need line-level placement, including
  // wrapped lines and multi-paragraph content. Preserve them for round trips,
  // but do not approximate them with block-level flex spacing.
  if (block.verticalAlign === "middle" || block.verticalAlign === "bottom") {
    containerEl.style.display = "flex";
    containerEl.style.flexDirection = "column";
    containerEl.style.justifyContent = block.verticalAlign === "middle" ? "center" : "flex-end";
  }

  // Store metadata
  containerEl.dataset["blockId"] = String(fragment.blockId);
  if (fragment.pmStart !== undefined) {
    containerEl.dataset["pmStart"] = String(fragment.pmStart);
  }
  if (fragment.pmEnd !== undefined) {
    containerEl.dataset["pmEnd"] = String(fragment.pmEnd);
  }

  // Render inner content using pre-measured data
  const innerWidth = fragment.width - margins.left - margins.right;
  const contentLayout = layoutTextBoxContent(block.content, measure.innerMeasures);

  for (let i = 0; i < block.content.length; i++) {
    const contentBlock = block.content[i];
    const contentMeasure = measure.innerMeasures[i];
    const placement = contentLayout.placements[i];
    if (!contentBlock || !contentMeasure || !placement) {
      continue;
    }

    if (contentBlock.kind === "table" && contentMeasure.kind === "table") {
      if (!options.renderTable) {
        panic("renderTextBoxFragment: a nested table renderer is required for table content");
      }
      const tableEl = options.renderTable(contentBlock, contentMeasure, context, doc, innerWidth);
      tableEl.style.marginTop = `${placement.leadingSpacing}px`;
      containerEl.append(tableEl);
      continue;
    }

    if (contentBlock.kind !== "paragraph" || contentMeasure.kind !== "paragraph") {
      continue;
    }

    const paraFragment = {
      kind: "paragraph" as const,
      blockId: contentBlock.id,
      x: 0,
      y: 0,
      width: innerWidth,
      height: placement.contentHeight,
      ...(contentBlock.pmStart !== undefined ? { pmStart: contentBlock.pmStart } : {}),
      ...(contentBlock.pmEnd !== undefined ? { pmEnd: contentBlock.pmEnd } : {}),
      fromLine: 0,
      toLine: contentMeasure.lines.length,
    };

    const previousBlock = block.content[i - 1];
    const nextBlock = block.content[i + 1];
    const prevBorders =
      previousBlock?.kind === "paragraph" ? previousBlock.attrs?.borders : undefined;
    const nextBorders = nextBlock?.kind === "paragraph" ? nextBlock.attrs?.borders : undefined;
    const paraEl = renderParagraphFragment(paraFragment, contentBlock, contentMeasure, context, {
      document: doc,
      ...(prevBorders !== undefined ? { prevBorders } : {}),
      ...(nextBorders !== undefined ? { nextBorders } : {}),
    });

    // Override absolute positioning to use relative flow within the text box
    paraEl.style.position = "relative";
    paraEl.style.left = "0";
    paraEl.style.top = "0";
    paraEl.style.height = `${placement.contentHeight}px`;
    paraEl.style.marginTop = `${placement.leadingSpacing}px`;

    containerEl.append(paraEl);
  }

  return containerEl;
}
