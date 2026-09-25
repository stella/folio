/**
 * Image conversion: inline image runs, block-level image blocks, and the
 * page-size constraint applied to both.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { ImageBlock, ImageRun, RunFormatting } from "../../layout-engine/types";
import { expectImageAttrs } from "../../prosemirror/attrs";
import type { ImageAttrs } from "../../prosemirror/schema/nodes";
import { cssBorderStyleOf } from "../../utils/borderCss";
import { sanitizeImageSrc } from "../../utils/sanitizeImageSrc";
import { nextBlockId } from "./flowConversionShared";

/**
 * Constrain image dimensions to fit within the page content area.
 * Scales proportionally if height exceeds pageContentHeight.
 */
export function constrainImageToPage(
  width: number,
  height: number,
  pageContentHeight: number | undefined,
): { width: number; height: number } {
  if (!pageContentHeight || height <= pageContentHeight) {
    return { width, height };
  }
  const scale = pageContentHeight / height;
  return { width: Math.round(width * scale), height: pageContentHeight };
}

/**
 * Build an ImageRun from ProseMirror node attrs, applying conditional property assignment
 * to satisfy exactOptionalPropertyTypes.
 */
export function buildImageRun(
  attrs: ImageAttrs,
  constrained: { width: number; height: number },
  pmStart: number,
  pmEnd: number,
  // Tracked-change attrs lifted off the image node's PM marks. eigenpal #641.
  trackedChange?: Pick<
    RunFormatting,
    "isInsertion" | "isDeletion" | "changeAuthor" | "changeDate" | "changeRevisionId"
  >,
  // Compatibility mode 15+ ignores an authored `layoutInCell="0"`; see
  // `resolveAnchorLayoutInCellCompatibility`.
  forceLayoutInCell?: boolean,
): ImageRun {
  const run: ImageRun = {
    kind: "image",
    src: attrs.src,
    ...(attrs.preview ? { preview: attrs.preview } : {}),
    width: constrained.width,
    height: constrained.height,
    pmStart,
    pmEnd,
  };
  if (attrs.alt !== undefined) {
    run.alt = attrs.alt;
  }
  if (attrs.transform !== undefined) {
    run.transform = attrs.transform;
  }
  // eigenpal #424 (opacity render pipeline): copy opacity verbatim. PM
  // schema defaults `opacity` to `null`, which survives the typed cast on
  // ImageAttrs (`number | undefined`). Gate with `!= null` so the model
  // never carries the schema sentinel.
  if (attrs.opacity != null) {
    run.opacity = attrs.opacity;
  }
  if (attrs.brightness != null) {
    run.brightness = attrs.brightness;
  }
  if (attrs.contrast != null) {
    run.contrast = attrs.contrast;
  }
  if (attrs.wrapType !== undefined) {
    run.wrapType = attrs.wrapType;
  }
  if (attrs.displayMode !== undefined) {
    run.displayMode = attrs.displayMode;
  }
  if (attrs.cssFloat !== undefined) {
    run.cssFloat = attrs.cssFloat;
  }
  if (attrs.distTop !== undefined) {
    run.distTop = attrs.distTop;
  }
  if (attrs.distBottom !== undefined) {
    run.distBottom = attrs.distBottom;
  }
  if (attrs.distLeft !== undefined) {
    run.distLeft = attrs.distLeft;
  }
  if (attrs.distRight !== undefined) {
    run.distRight = attrs.distRight;
  }
  if (attrs._docxObjectPreview === true) {
    run.exactLineHeight = true;
  }
  // eigenpal #424: pass crop fractions through to the painter so it can
  // emit CSS clip-path. PM defaults are `null`; treat null as "not set".
  if (attrs.cropTop != null) {
    run.cropTop = attrs.cropTop;
  }
  if (attrs.cropRight != null) {
    run.cropRight = attrs.cropRight;
  }
  if (attrs.cropBottom != null) {
    run.cropBottom = attrs.cropBottom;
  }
  if (attrs.cropLeft != null) {
    run.cropLeft = attrs.cropLeft;
  }
  // eigenpal #1096: image borders are authored on the PM image attrs and
  // painted by layout-painter. PM defaults are null; treat null as absent.
  if (attrs.borderWidth != null) {
    run.borderWidth = attrs.borderWidth;
  }
  if (attrs.borderColor) {
    run.borderColor = attrs.borderColor;
  }
  const runBorderStyle = cssBorderStyleOf(attrs.borderStyle);
  if (runBorderStyle) {
    run.borderStyle = runBorderStyle;
  }
  if (attrs.position !== undefined) {
    run.position = attrs.position;
  }
  if (forceLayoutInCell) {
    run.layoutInCell = true;
  } else if (attrs.anchor?.layoutInCell !== undefined) {
    run.layoutInCell = attrs.anchor.layoutInCell;
  }
  if (trackedChange?.isInsertion) {
    run.isInsertion = true;
  }
  if (trackedChange?.isDeletion) {
    run.isDeletion = true;
  }
  if (trackedChange?.changeAuthor !== undefined) {
    run.changeAuthor = trackedChange.changeAuthor;
  }
  if (trackedChange?.changeDate !== undefined) {
    run.changeDate = trackedChange.changeDate;
  }
  if (trackedChange?.changeRevisionId !== undefined) {
    run.changeRevisionId = trackedChange.changeRevisionId;
  }
  return run;
}

/** A package image whose bytes cannot paint still owns its authored line box. */
export const hasRelationshipBackedImageBox = (attrs: ImageAttrs): boolean =>
  typeof attrs.rId === "string" &&
  attrs.rId.trim().length > 0 &&
  typeof attrs.width === "number" &&
  Number.isFinite(attrs.width) &&
  attrs.width >= 0 &&
  typeof attrs.height === "number" &&
  Number.isFinite(attrs.height) &&
  attrs.height >= 0;

/**
 * A preview counts as a source: the bytes are not in the attrs, but the
 * display list can produce them without I/O when it interns the descriptor.
 */
export const hasPaintableImageSource = (attrs: Pick<ImageAttrs, "src" | "preview">): boolean =>
  attrs.preview !== undefined || sanitizeImageSrc(attrs.src) !== undefined;

/**
 * Convert an image node to an ImageBlock.
 */
export function convertImage(
  node: PMNode,
  startPos: number,
  pageContentHeight?: number,
): ImageBlock | undefined {
  const attrs = expectImageAttrs(node);
  if (!hasPaintableImageSource(attrs) && !hasRelationshipBackedImageBox(attrs)) {
    return undefined;
  }
  const wrapType = attrs.wrapType;

  // Only anchor images with 'behind' or 'inFront' wrap types
  // Other wrap types (square, tight, through, topAndBottom) need text wrapping
  // which we don't support yet, so treat them as block-level images
  const shouldAnchor = wrapType === "behind" || wrapType === "inFront";

  const constrained = constrainImageToPage(
    attrs.width ?? 100,
    attrs.height ?? 100,
    pageContentHeight,
  );

  const imgBlock: ImageBlock = {
    kind: "image",
    id: nextBlockId(),
    src: attrs.src,
    ...(attrs.preview ? { preview: attrs.preview } : {}),
    width: constrained.width,
    height: constrained.height,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
  if (attrs.alt) {
    imgBlock.alt = attrs.alt;
  }
  if (attrs.transform) {
    imgBlock.transform = attrs.transform;
  }
  // eigenpal #424 (opacity render pipeline). `!= null` so PM's null schema
  // default doesn't leak into ImageBlock.opacity (`number | undefined`).
  if (attrs.opacity != null) {
    imgBlock.opacity = attrs.opacity;
  }
  if (attrs.brightness != null) {
    imgBlock.brightness = attrs.brightness;
  }
  if (attrs.contrast != null) {
    imgBlock.contrast = attrs.contrast;
  }
  if (shouldAnchor) {
    const anchor: NonNullable<ImageBlock["anchor"]> = {
      isAnchored: true,
      behindDoc: wrapType === "behind",
    };
    if (attrs.distLeft !== undefined) {
      anchor.offsetH = attrs.distLeft;
    }
    if (attrs.distTop !== undefined) {
      anchor.offsetV = attrs.distTop;
    }
    imgBlock.anchor = anchor;
  }
  if (attrs.hlinkHref) {
    imgBlock.hlinkHref = attrs.hlinkHref;
  }
  // eigenpal #424: thread wp:srcRect crop fractions to the floating-image
  // block so renderers can apply clip-path consistently across paths.
  if (attrs.cropTop != null) {
    imgBlock.cropTop = attrs.cropTop;
  }
  if (attrs.cropRight != null) {
    imgBlock.cropRight = attrs.cropRight;
  }
  if (attrs.cropBottom != null) {
    imgBlock.cropBottom = attrs.cropBottom;
  }
  if (attrs.cropLeft != null) {
    imgBlock.cropLeft = attrs.cropLeft;
  }
  // eigenpal #1096: preserve image border attrs for floating/block image
  // painting. PM defaults are null; treat null as absent.
  if (attrs.borderWidth != null) {
    imgBlock.borderWidth = attrs.borderWidth;
  }
  if (attrs.borderColor) {
    imgBlock.borderColor = attrs.borderColor;
  }
  const imageBorderStyle = cssBorderStyleOf(attrs.borderStyle);
  if (imageBorderStyle) {
    imgBlock.borderStyle = imageBorderStyle;
  }
  return imgBlock;
}
