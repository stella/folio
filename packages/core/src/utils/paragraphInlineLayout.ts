import { panic } from "better-result";

import { getListMarkerInlineWidth } from "../layout-engine/measure/listMarkerWidth";
import type {
  ImageRun,
  MeasuredLine,
  Page,
  ParagraphAttrs,
  ParagraphBlock,
  ParagraphFragment,
} from "../layout-engine/types";
import { isFloatingImageRun } from "../layout-engine/types";
import { isRtlParagraph } from "./paragraphBaseDirection";
import { getParagraphFrame } from "./paragraphFrame";
import { inlineImageBoundingBox } from "./rotationBoundingBox";

type ParagraphAlignment = NonNullable<ParagraphAttrs["alignment"]>;

export type HorizontalClipViewport = {
  left: number;
  right: number;
};

export const intersectHorizontalClipViewports = (
  first: HorizontalClipViewport | undefined,
  second: HorizontalClipViewport,
): HorizontalClipViewport => {
  if (first === undefined) return second;
  const left = Math.max(first.left, second.left);
  return { left, right: Math.max(left, Math.min(first.right, second.right)) };
};

export type InlinePictureClipInsets = {
  left: number;
  right: number;
};

type HorizontalColumn = {
  left: number;
  width: number;
};

type ColumnPictureClipViewportOptions = {
  paper: HorizontalClipViewport;
  columns: readonly HorizontalColumn[];
  columnIndex: number;
};

export type PageColumnLayout = NonNullable<Page["columns"]>;

export type InlinePictureVisibleBounds = {
  left: number;
  right: number;
};

type InlinePictureVisibleBoundsOptions = {
  pictureLeft: number;
  pictureWidth: number;
  viewport: HorizontalClipViewport;
};

export const inlinePictureVisibleBounds = ({
  pictureLeft,
  pictureWidth,
  viewport,
}: InlinePictureVisibleBoundsOptions): InlinePictureVisibleBounds => {
  const left = Math.max(pictureLeft, viewport.left);
  return {
    left,
    right: Math.max(left, Math.min(pictureLeft + pictureWidth, viewport.right)),
  };
};

type PageHorizontalColumnsOptions = {
  layout: PageColumnLayout | undefined;
  contentLeft: number;
  contentWidth: number;
};

export const pageHorizontalColumns = ({
  layout,
  contentLeft,
  contentWidth,
}: PageHorizontalColumnsOptions): HorizontalColumn[] => {
  if (layout === undefined || layout.count < 2) {
    return [];
  }
  const equalWidth =
    (contentWidth - layout.gap * (layout.count - 1)) / layout.count;
  const columns: HorizontalColumn[] = [];
  let left = contentLeft;
  for (let index = 0; index < layout.count; index += 1) {
    const width = layout.widths?.at(index) ?? equalWidth;
    columns.push({ left, width });
    left += width + (layout.gaps?.[index] ?? layout.gap);
  }
  return columns;
};

type ColumnPictureClipViewportForFragmentOptions = {
  paper: HorizontalClipViewport;
  columns: readonly HorizontalColumn[];
  fragmentX: number;
  fragmentWidth: number;
  inlinePictureViewport?: HorizontalClipViewport;
};

export const columnPictureClipViewportForFragment = ({
  paper,
  columns,
  fragmentX,
  fragmentWidth,
  inlinePictureViewport,
}: ColumnPictureClipViewportForFragmentOptions): HorizontalClipViewport => {
  if (inlinePictureViewport !== undefined) return inlinePictureViewport;
  if (columns.length === 0) {
    return paper;
  }
  const columnIndex = columns.findIndex(
    (column) => column.left === fragmentX && column.width === fragmentWidth,
  );
  if (columnIndex === -1) {
    return panic("Paragraph fragment has no matching page column clip bounds");
  }
  return columnPictureClipViewportForColumn({ paper, columns, columnIndex });
};

type InlinePictureClipInsetsOptions = {
  pictureLeft: number;
  pictureWidth: number;
  viewport: HorizontalClipViewport;
};

export const inlinePictureClipInsets = ({
  pictureLeft,
  pictureWidth,
  viewport,
}: InlinePictureClipInsetsOptions): InlinePictureClipInsets => ({
  left: Math.max(0, viewport.left - pictureLeft),
  right: Math.max(0, pictureLeft + pictureWidth - viewport.right),
});

export const columnPictureClipViewportForColumn = ({
  paper,
  columns,
  columnIndex,
}: ColumnPictureClipViewportOptions): HorizontalClipViewport => {
  const column = columns[columnIndex];
  if (column === undefined || columns.length < 2) {
    return paper;
  }

  const columnRight = column.left + column.width;
  let left = paper.left;
  let right = paper.right;
  for (const [index, other] of columns.entries()) {
    if (index === columnIndex) {
      continue;
    }
    const otherRight = other.left + other.width;
    if (otherRight <= column.left) {
      left = Math.max(left, (otherRight + column.left) / 2);
    }
    if (other.left >= columnRight) {
      right = Math.min(right, (columnRight + other.left) / 2);
    }
  }

  return { left, right };
};

export const resolveImageLineAlign = (
  imageRun: ImageRun,
  paragraphAlignment: ParagraphAlignment | undefined,
): ParagraphAlignment | undefined => {
  const horizontal = imageRun.position?.horizontal;
  if (!horizontal) {
    return paragraphAlignment;
  }
  switch (horizontal.align) {
    case "center":
      return "center";
    case "right":
    case "outside":
      return "right";
    default:
      return "left";
  }
};

type ResolveInlineImageLineAlignOptions = {
  imageRun: ImageRun;
  paragraphAlignment: ParagraphAlignment | undefined;
  availableWidth: number | undefined;
};

const resolveInlineImageLineAlign = ({
  imageRun,
  paragraphAlignment,
  availableWidth,
}: ResolveInlineImageLineAlignOptions): ParagraphAlignment | undefined => {
  if (
    !imageRun.position?.horizontal &&
    availableWidth !== undefined &&
    inlineImageBoundingBox(imageRun).width > availableWidth
  ) {
    return "left";
  }
  return resolveImageLineAlign(imageRun, paragraphAlignment);
};

type ResolveParagraphLineAlignmentOptions = {
  imageOnlyRun?: ImageRun;
  paragraphAlignment: ParagraphAlignment;
  availableWidth: number;
};

export const resolveParagraphLineAlignment = ({
  imageOnlyRun,
  paragraphAlignment,
  availableWidth,
}: ResolveParagraphLineAlignmentOptions): ParagraphAlignment =>
  imageOnlyRun === undefined
    ? paragraphAlignment
    : (resolveInlineImageLineAlign({
        imageRun: imageOnlyRun,
        paragraphAlignment,
        availableWidth,
      }) ?? paragraphAlignment);

export const paragraphImageOnlyRun = (
  block: ParagraphBlock,
  line: MeasuredLine,
): ImageRun | undefined => {
  if (line.fromRun !== line.toRun) return undefined;
  const run = block.runs.at(line.fromRun);
  if (run?.kind !== "image" || !isInlinePictureRun(run)) return undefined;
  return run;
};

export const paragraphHasInlinePicture = (block: ParagraphBlock): boolean =>
  block.runs.some((run) => run.kind === "image" && isInlinePictureRun(run));

type ParagraphInlinePictureViewportOptions = {
  block: ParagraphBlock;
  fragment: Pick<ParagraphFragment, "x" | "width" | "inlinePictureViewport">;
  paper: HorizontalClipViewport;
  columns: readonly HorizontalColumn[];
};

export const paragraphInlinePictureViewport = ({
  block,
  fragment,
  paper,
  columns,
}: ParagraphInlinePictureViewportOptions):
  HorizontalClipViewport | undefined => {
  if (!paragraphHasInlinePicture(block)) return undefined;
  if (getParagraphFrame(block) !== undefined) {
    return { left: fragment.x, right: fragment.x + fragment.width };
  }
  return columnPictureClipViewportForFragment({
    paper,
    columns,
    fragmentX: fragment.x,
    fragmentWidth: fragment.width,
    ...(fragment.inlinePictureViewport === undefined
      ? {}
      : { inlinePictureViewport: fragment.inlinePictureViewport }),
  });
};

export const isInlinePictureRun = (run: ImageRun): boolean =>
  !isFloatingImageRun(run) &&
  run.displayMode !== "block" &&
  run.wrapType !== "topAndBottom";

export type ParagraphLineGeometry = {
  alignment: ParagraphAlignment;
  paintStartXPx: number;
  layoutStartXPx: number;
  layoutRightEdgeXPx: number;
  baselineYPx: number;
  lineTopYPx: number;
};

type ResolveParagraphLineGeometryOptions = {
  fragment: Pick<ParagraphFragment, "x" | "width">;
  line: MeasuredLine;
  lineTopYPx: number;
  alignment: ParagraphAlignment;
  isRtl: boolean;
  indentLeft: number;
  indentRight: number;
  firstLineOffsetPx: number;
  markerInlineWidthPx: number;
  isFirstLine: boolean;
  imageOnlyRun?: ImageRun;
};

export const resolveParagraphLineGeometry = ({
  fragment,
  line,
  lineTopYPx,
  alignment,
  isRtl,
  indentLeft,
  indentRight,
  firstLineOffsetPx,
  markerInlineWidthPx,
  isFirstLine,
  imageOnlyRun,
}: ResolveParagraphLineGeometryOptions): ParagraphLineGeometry => {
  const leftOffset = line.leftOffset ?? 0;
  const rightOffset = line.rightOffset ?? 0;
  const firstLineExtra = isFirstLine
    ? firstLineOffsetPx + markerInlineWidthPx
    : 0;
  const layoutStartXPx = indentLeft + leftOffset + (isRtl ? 0 : firstLineExtra);
  const layoutRightEdgeXPx = fragment.width - indentRight - rightOffset;
  const leftEdge = fragment.x + layoutStartXPx;
  const rightEdge =
    fragment.x + layoutRightEdgeXPx - (isRtl ? firstLineExtra : 0);
  const resolvedAlignment = resolveParagraphLineAlignment({
    ...(imageOnlyRun === undefined ? {} : { imageOnlyRun }),
    paragraphAlignment: alignment,
    availableWidth: rightEdge - leftEdge,
  });

  let paintStartXPx: number;
  switch (resolvedAlignment) {
    case "right":
      paintStartXPx = rightEdge - line.width;
      break;
    case "center":
      paintStartXPx = leftEdge + (rightEdge - leftEdge - line.width) / 2;
      break;
    case "left":
    case "justify":
      paintStartXPx = isRtl ? rightEdge - line.width : leftEdge;
      break;
  }

  const baselineYPx =
    lineTopYPx +
    (line.lineHeight - (line.ascent + line.descent)) / 2 +
    line.ascent;

  return {
    alignment: resolvedAlignment,
    paintStartXPx,
    layoutStartXPx,
    layoutRightEdgeXPx,
    baselineYPx,
    lineTopYPx,
  };
};

type ResolveParagraphLineGeometryForBlockOptions = {
  fragment: Pick<ParagraphFragment, "x" | "width">;
  block: ParagraphBlock;
  line: MeasuredLine;
  lineTopYPx: number;
  isFirstLine: boolean;
};

export const resolveParagraphLineGeometryForBlock = ({
  fragment,
  block,
  line,
  lineTopYPx,
  isFirstLine,
}: ResolveParagraphLineGeometryForBlockOptions): ParagraphLineGeometry => {
  const inlineLayout = resolvePhysicalParagraphInlineLayout(block);
  const indent = block.attrs?.indent;
  const firstLineOffsetPx =
    indent?.hanging !== undefined && indent.hanging > 0
      ? -indent.hanging
      : (indent?.firstLine ?? 0);
  const imageOnlyRun = paragraphImageOnlyRun(block, line);
  return resolveParagraphLineGeometry({
    fragment,
    line,
    lineTopYPx,
    alignment: inlineLayout.alignment,
    isRtl: inlineLayout.isRtl,
    indentLeft: inlineLayout.indentLeft,
    indentRight: inlineLayout.indentRight,
    firstLineOffsetPx,
    markerInlineWidthPx: getListMarkerInlineWidth(block),
    isFirstLine,
    ...(imageOnlyRun === undefined ? {} : { imageOnlyRun }),
  });
};

export type PhysicalParagraphInlineLayout = {
  alignment: ParagraphAlignment;
  explicitAlignment?: ParagraphAlignment;
  indentLeft: number;
  indentRight: number;
  isRtl: boolean;
};

const mirrorHorizontalAlignment = (
  alignment: ParagraphAlignment | undefined,
): ParagraphAlignment | undefined => {
  if (alignment === "left") {
    return "right";
  }
  if (alignment === "right") {
    return "left";
  }
  return alignment;
};

/** Resolve logical OOXML paragraph sides into physical inline geometry. */
export const resolvePhysicalParagraphInlineLayout = (
  block: ParagraphBlock,
): PhysicalParagraphInlineLayout => {
  const attrs = block.attrs;
  const isRtl = isRtlParagraph(block);
  const mirrorsHorizontalSides = attrs?.bidi === true;
  const explicitAlignment = mirrorsHorizontalSides
    ? mirrorHorizontalAlignment(attrs?.alignment)
    : attrs?.alignment;
  const indentLeft = attrs?.indent?.left ?? 0;
  const indentRight = attrs?.indent?.right ?? 0;

  return {
    alignment: explicitAlignment ?? (isRtl ? "right" : "left"),
    ...(explicitAlignment === undefined ? {} : { explicitAlignment }),
    indentLeft: mirrorsHorizontalSides ? indentRight : indentLeft,
    indentRight: mirrorsHorizontalSides ? indentLeft : indentRight,
    isRtl,
  };
};
