/**
 * Text Box Renderer
 *
 * Renders text box fragments to DOM. Handles:
 * - Background fill color
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
  TextBoxMeasure,
} from "../layout-engine/types";
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
  ) => HTMLElement;
};

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
  containerEl.style.overflow = "hidden";
  containerEl.style.boxSizing = "border-box";

  // Fill color
  if (block.fillColor) {
    containerEl.style.backgroundColor = block.fillColor;
  }

  // Border/outline
  if (block.outlineWidth && block.outlineWidth > 0) {
    const style = block.outlineStyle || "solid";
    const color = block.outlineColor || "#000000";
    containerEl.style.border = `${block.outlineWidth}px ${style} ${color}`;
  }

  // Internal padding
  const margins = block.margins ?? DEFAULT_TEXTBOX_MARGINS;
  containerEl.style.padding = `${margins.top}px ${margins.right}px ${margins.bottom}px ${margins.left}px`;

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
      const tableEl = options.renderTable(contentBlock, contentMeasure, context, doc);
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
