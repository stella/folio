/**
 * Stacked story blocks: the header, the footer and a footnote body.
 *
 * These three are not fragments. The layout engine paginates the body only, so
 * a story arrives as a `(blocks, measures)` pair already measured against its
 * own width, and whoever paints it stacks the blocks itself. Every block is
 * then handed to the *same* paragraph, table, image and text-box builders the
 * body uses: a second text painter for headers is exactly how a header comes to
 * disagree with the body about a tab stop or a justified space.
 *
 * The stacking rules differ between the two callers and both are the painter's:
 * `renderFootnoteContent` advances by the measured height of every block, while
 * `renderHeaderFooterContent` honours explicit `w:spacing w:before`, offsets an
 * inline table by `w:jc`/`w:tblInd`, and lets a floating or page-anchored
 * object out of the flow entirely.
 */

import { isFloatingTextBoxBlock } from "../../layout-engine/types";
import type {
  FlowBlock,
  ImageBlock,
  Measure,
  ParagraphBlock,
  TableBlock,
  TextBoxBlock,
} from "../../layout-engine/types";
import { isFloatingImageRun } from "../../layout-painter/renderUtils";
import type { BuildContext } from "./buildContext";
import { blockRegion, type PageComposer } from "./regions";
import { HIT_REGION_KINDS } from "../primitives";
import { paintImageFragment } from "./imagePrimitives";
import { paintParagraphFragment } from "./paragraphPrimitives";
import { paintTableFragment } from "./tablePrimitives";
import { paintTextBoxFragment } from "./textBoxPrimitives";
import { UNSUPPORTED_CONSTRUCT, type UnsupportedConstruct } from "./unsupported";

/** Mirrors `isPositionedHeaderFooterTextBoxBlock`: an explicit display mode keeps a box in the flow. */
const isPositionedTextBox = (block: TextBoxBlock): boolean =>
  block.displayMode !== "block" && block.displayMode !== "inline" && isFloatingTextBoxBlock(block);

const isAnchoredImage = (block: ImageBlock): boolean => block.anchor?.isAnchored === true;

/** `w:jc` / `w:tblInd` for an inline story table, as the body's `desiredX` resolves them. */
const inlineTableOffsetPx = (block: TableBlock, tableWidthPx: number, widthPx: number): number => {
  if (block.justification === "center") {
    return (widthPx - tableWidthPx) / 2;
  }
  if (block.justification === "right") {
    return widthPx - tableWidthPx;
  }
  return block.indent ?? 0;
};

/**
 * Explicit `w:spacing w:before` only. `normalizeHeaderFooterMeasureBlocks`
 * strips style-inherited spacing from the measurement copy, so offsetting by an
 * inherited value would push the first line below the space its own measure
 * reserved and every following block with it.
 */
const explicitSpaceBeforePx = (block: ParagraphBlock): number =>
  block.attrs?.spacingExplicit?.before === true ? (block.attrs.spacing?.before ?? 0) : 0;

type PaintStoryBlockOptions = {
  readonly composer: PageComposer;
  readonly block: FlowBlock;
  readonly measure: Measure;
  readonly xPx: number;
  readonly yPx: number;
  readonly widthPx: number;
  readonly context: BuildContext;
  /** Names the story in an `unsupported` entry, e.g. `"footnote 3"`. */
  readonly label: string;
  readonly construct: UnsupportedConstruct;
};

/** One block of a story, painted at an absolute page position. */
const paintStoryBlock = ({
  composer,
  block,
  measure,
  xPx,
  yPx,
  widthPx,
  context,
  label,
  construct,
}: PaintStoryBlockOptions): void => {
  if (block.kind === "paragraph" && measure.kind === "paragraph") {
    const fragment = {
      kind: "paragraph",
      blockId: block.id,
      x: xPx,
      y: yPx,
      width: widthPx,
      height: measure.totalHeight,
      fromLine: 0,
      toLine: measure.lines.length,
    } as const;
    composer.region(blockRegion({ fragment, kind: HIT_REGION_KINDS.paragraph, context }), () => {
      paintParagraphFragment({ fragment, block, measure, context, composer });
    });
    return;
  }

  if (block.kind === "table" && measure.kind === "table") {
    const fragment = {
      kind: "table",
      blockId: block.id,
      x: xPx,
      y: yPx,
      width: measure.totalWidth,
      height: measure.totalHeight,
      fromRow: 0,
      toRow: block.rows.length,
    } as const;
    composer.region(blockRegion({ fragment, kind: HIT_REGION_KINDS.table, context }), () => {
      paintTableFragment({ fragment, block, measure, context, composer });
    });
    return;
  }

  if (block.kind === "image" && measure.kind === "image") {
    const fragment = {
      kind: "image",
      blockId: block.id,
      x: xPx,
      y: yPx,
      width: measure.width,
      height: measure.height,
    } as const;
    composer.region(blockRegion({ fragment, kind: HIT_REGION_KINDS.image, context }), () => {
      composer.push(paintImageFragment(fragment, block, context));
    });
    return;
  }

  if (block.kind === "textBox" && measure.kind === "textBox") {
    const fragment = {
      kind: "textBox",
      blockId: block.id,
      x: xPx,
      y: yPx,
      width: measure.width,
      height: measure.height,
    } as const;
    composer.region(blockRegion({ fragment, kind: HIT_REGION_KINDS.textBox, context }), () => {
      paintTextBoxFragment({ composer, fragment, block, measure, context });
    });
    return;
  }

  context.unsupported.report(
    construct,
    context.pageIndex,
    `${label}: a ${block.kind} block with a ${measure.kind} measure is not painted`,
  );
};

const measuredHeightPx = (measure: Measure): number => {
  switch (measure.kind) {
    case "paragraph":
    case "table":
      return measure.totalHeight;
    case "image":
    case "textBox":
      return measure.height;
    case "sectionBreak":
    case "pageBreak":
    case "columnBreak":
      return 0;
    default:
      measure satisfies never;
      return 0;
  }
};

export type PaintStoryOptions = {
  readonly composer: PageComposer;
  readonly blocks: readonly FlowBlock[];
  readonly measures: readonly Measure[];
  /** Page-absolute origin of the story's own content box. */
  readonly xPx: number;
  readonly yPx: number;
  readonly widthPx: number;
  readonly context: BuildContext;
  readonly label: string;
};

/**
 * A footnote body: every block stacks by its measured height, which is the
 * height the paginator reserved the band from.
 */
export const paintFootnoteBlocks = ({
  composer,
  blocks,
  measures,
  xPx,
  yPx,
  widthPx,
  context,
  label,
}: PaintStoryOptions): void => {
  let cursorYPx = yPx;

  for (const [index, block] of blocks.entries()) {
    const measure = measures[index];
    if (measure === undefined) {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.missingBlock,
        context.pageIndex,
        `${label}: block ${String(block.id)} has no measure`,
      );
      continue;
    }
    paintStoryBlock({
      composer,
      block,
      measure,
      xPx,
      yPx: cursorYPx,
      widthPx,
      context,
      label,
      construct: UNSUPPORTED_CONSTRUCT.footnoteContent,
    });
    cursorYPx += measuredHeightPx(measure);
  }
};

/**
 * A header or footer story. Blocks that Word takes out of the flow (a floating
 * table, an anchored image, a positioned text box) contribute no height, and
 * they are reported rather than painted: their position resolves against page
 * and margin anchors the display list producer does not carry.
 */
export const paintHeaderFooterBlocks = ({
  composer,
  blocks,
  measures,
  xPx,
  yPx,
  widthPx,
  context,
  label,
}: PaintStoryOptions): void => {
  let cursorYPx = yPx;

  const reportOutOfFlow = (detail: string): void => {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.headerFooterContent,
      context.pageIndex,
      `${label}: ${detail}`,
    );
  };

  for (const [index, block] of blocks.entries()) {
    const measure = measures[index];
    if (measure === undefined) {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.missingBlock,
        context.pageIndex,
        `${label}: block ${String(block.id)} has no measure`,
      );
      continue;
    }

    switch (block.kind) {
      case "paragraph": {
        if (measure.kind !== "paragraph") {
          break;
        }
        if (
          block.runs.some(
            (run) =>
              run.kind === "image" && (isFloatingImageRun(run) || run.position !== undefined),
          )
        ) {
          reportOutOfFlow(
            `paragraph ${String(block.id)} anchors a floating picture, which resolves against page and margin anchors the display list does not carry`,
          );
        }
        paintStoryBlock({
          composer,
          block,
          measure,
          xPx,
          yPx: cursorYPx + explicitSpaceBeforePx(block),
          widthPx,
          context,
          label,
          construct: UNSUPPORTED_CONSTRUCT.headerFooterContent,
        });
        cursorYPx += measure.totalHeight;
        continue;
      }
      case "table": {
        if (measure.kind !== "table") {
          break;
        }
        if (block.floating) {
          reportOutOfFlow(`table ${String(block.id)} is floating (w:tblpPr) and is not painted`);
          continue;
        }
        paintStoryBlock({
          composer,
          block,
          measure,
          xPx: xPx + inlineTableOffsetPx(block, measure.totalWidth, widthPx),
          yPx: cursorYPx,
          widthPx,
          context,
          label,
          construct: UNSUPPORTED_CONSTRUCT.headerFooterContent,
        });
        cursorYPx += measure.totalHeight;
        continue;
      }
      case "image": {
        if (measure.kind !== "image") {
          break;
        }
        if (isAnchoredImage(block)) {
          reportOutOfFlow(`image ${String(block.id)} is anchored and is not painted`);
          continue;
        }
        paintStoryBlock({
          composer,
          block,
          measure,
          xPx,
          yPx: cursorYPx,
          widthPx,
          context,
          label,
          construct: UNSUPPORTED_CONSTRUCT.headerFooterContent,
        });
        cursorYPx += measure.height;
        continue;
      }
      case "textBox": {
        if (measure.kind !== "textBox") {
          break;
        }
        if (isPositionedTextBox(block)) {
          reportOutOfFlow(`text box ${String(block.id)} is positioned and is not painted`);
          continue;
        }
        paintStoryBlock({
          composer,
          block,
          measure,
          xPx,
          yPx: cursorYPx,
          widthPx,
          context,
          label,
          construct: UNSUPPORTED_CONSTRUCT.headerFooterContent,
        });
        cursorYPx += measure.height;
        continue;
      }
      case "pageBreak":
      case "columnBreak":
      case "sectionBreak":
        // A break inside a header or footer paints nothing and reserves
        // nothing: the story is placed on every page, not paginated.
        continue;
      default:
        block satisfies never;
        continue;
    }

    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.measureMismatch,
      context.pageIndex,
      `${label}: block ${String(block.id)} is a ${block.kind} with a ${measure.kind} measure`,
    );
  }
};
