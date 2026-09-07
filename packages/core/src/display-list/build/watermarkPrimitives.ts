/**
 * The document watermark, behind the page's content.
 *
 * Word stores it as a VML shape in a header part; the painter reproduces it
 * with CSS rather than mirroring `v:textpath`, and this reproduces the same
 * result as primitives: a text watermark is one rotated, translucent run
 * centred on the page; a picture watermark is the image scaled to fit a
 * fraction of the page and washed out.
 *
 * A picture watermark needs its bytes. The relationship id → asset mapping
 * belongs to the package layer, so a caller that cannot resolve it hands over
 * no source and the watermark is reported rather than dropped: a page missing
 * its watermark and a page that never had one must not look alike.
 */

import { ptToPx, pxToPt } from "../../layout-engine/measure/measureHelpers";
import { getFontMetrics } from "../../layout-engine/measure/measureProvider";
import type { FontStyle } from "../../layout-engine/measure/measureTypes";
import type { Page } from "../../layout-engine/types";
import type { Watermark } from "../../types/document";
import type { DisplayImagePrimitive, DisplayPrimitive, DisplayRect } from "../types";
import type { BuildContext } from "./buildContext";
import { parseDisplayColor } from "./colors";
import { buildGlyphs, glyphRunText } from "./glyphs";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

/**
 * `renderWatermark.ts`: Word's text-watermark preset is ~144px bold silver at
 * half opacity, rotated bottom-left to top-right.
 */
const TEXT_FONT_SIZE_PX = 144;
const TEXT_DEFAULT_FONT = "Calibri";
const TEXT_DEFAULT_COLOR = "#C0C0C0";
const TEXT_DEFAULT_OPACITY = 0.5;
const TEXT_DIAGONAL_DEGREES = -45;
const PICTURE_NATIVE_SCALE = 1;
const PICTURE_WASHOUT_OPACITY = 0.18;

const paintTextWatermark = (
  watermark: Extract<Watermark, { kind: "text" }>,
  page: Page,
  context: BuildContext,
): readonly DisplayPrimitive[] => {
  if (watermark.text.length === 0) {
    return [];
  }

  const style: FontStyle = {
    fontFamily: watermark.font ?? TEXT_DEFAULT_FONT,
    fontSize: pxToPt(TEXT_FONT_SIZE_PX),
    bold: true,
  };
  const glyphs = buildGlyphs({
    text: watermark.text,
    style,
    allCaps: false,
    spaceDeltaPx: 0,
    collapsed: false,
  });

  // `color: "auto"` is a documented model value meaning "the producer
  // default", which for a watermark is silver rather than ink.
  const authored =
    watermark.color === undefined || watermark.color === "auto"
      ? TEXT_DEFAULT_COLOR
      : `#${watermark.color}`;
  const color = parseDisplayColor(authored);
  if (!color) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.unresolvedColor,
      context.pageIndex,
      `watermark colour ${authored}`,
    );
    return [];
  }

  // The painter centres the shape in the page box, so its own centre is the
  // page centre and the rotation turns about that point.
  const centerXPx = page.size.w / 2;
  const centerYPx = page.size.h / 2;
  const metrics = getFontMetrics(style);
  const baselineYPx =
    centerYPx - (metrics.fontBoxAscent + metrics.fontBoxDescent) / 2 + metrics.fontBoxAscent;

  return [
    {
      kind: "rotateGroup",
      degrees: watermark.diagonal === false ? 0 : TEXT_DIAGONAL_DEGREES,
      originXPx: centerXPx,
      originYPx: centerYPx,
      children: [
        {
          kind: "opacityGroup",
          opacity: watermark.opacity ?? TEXT_DEFAULT_OPACITY,
          children: [
            {
              kind: "glyphRun",
              font: context.fonts.intern({
                fontFamily: style.fontFamily ?? TEXT_DEFAULT_FONT,
                bold: true,
                measureStyle: style,
              }),
              fontSizePx: TEXT_FONT_SIZE_PX,
              color,
              xPx: centerXPx - glyphs.widthPx / 2,
              baselineYPx,
              ...glyphRunText(glyphs),
              direction: "ltr",
            },
          ],
        },
      ],
    },
  ];
};

/** `max-width`/`max-height` at `scale` of the page, `object-fit: contain`. */
const containedRect = (
  page: Page,
  scale: number,
  pixelWidth: number,
  pixelHeight: number,
): DisplayRect => {
  const boxWidthPx = page.size.w * scale;
  const boxHeightPx = page.size.h * scale;
  if (pixelWidth <= 0 || pixelHeight <= 0) {
    return { xPx: 0, yPx: 0, widthPx: boxWidthPx, heightPx: boxHeightPx };
  }
  const fit = Math.min(boxWidthPx / pixelWidth, boxHeightPx / pixelHeight);
  const widthPx = pixelWidth * fit;
  const heightPx = pixelHeight * fit;
  return {
    xPx: (page.size.w - widthPx) / 2,
    yPx: (page.size.h - heightPx) / 2,
    widthPx,
    heightPx,
  };
};

/** The centred VML shape box, when both authored dimensions survived parsing. */
const authoredPictureRect = (
  page: Page,
  widthPt: number | undefined,
  heightPt: number | undefined,
): DisplayRect | undefined => {
  if (widthPt === undefined || heightPt === undefined) {
    return undefined;
  }
  const widthPx = ptToPx(widthPt);
  const heightPx = ptToPx(heightPt);
  return {
    xPx: (page.size.w - widthPx) / 2,
    yPx: (page.size.h - heightPx) / 2,
    widthPx,
    heightPx,
  };
};

const paintPictureWatermark = (
  watermark: Extract<Watermark, { kind: "picture" }>,
  page: Page,
  imageSrc: string | undefined,
  context: BuildContext,
): readonly DisplayPrimitive[] => {
  if (imageSrc === undefined) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.watermark,
      context.pageIndex,
      `picture watermark ${watermark.imageRId} has no resolved image source: the relationship id resolves in the package layer, not in the builder`,
    );
    return [];
  }

  const ref = context.images.intern(imageSrc);
  if (ref === undefined) {
    const reason = context.images.failureFor(imageSrc) ?? "image could not be decoded";
    context.unsupported.report(
      reason.includes("not PNG or JPEG")
        ? UNSUPPORTED_CONSTRUCT.imageFormat
        : UNSUPPORTED_CONSTRUCT.imageSource,
      context.pageIndex,
      `picture watermark: ${reason}`,
    );
    return [];
  }

  const source = context.images.snapshot()[ref];
  const image: DisplayImagePrimitive = {
    kind: "image",
    image: ref,
    rect:
      authoredPictureRect(page, watermark.widthPt, watermark.heightPt) ??
      containedRect(
        page,
        watermark.scale ?? PICTURE_NATIVE_SCALE,
        source?.pixelWidth ?? 0,
        source?.pixelHeight ?? 0,
      ),
    opacity: watermark.washout === false ? 1 : PICTURE_WASHOUT_OPACITY,
  };
  return [image];
};

export type WatermarkPaintOptions = {
  readonly page: Page;
  readonly watermark: Watermark;
  /** Resolved `data:` source for a picture watermark; absent means unresolved. */
  readonly imageSrc?: string;
  readonly context: BuildContext;
};

export const paintWatermark = ({
  page,
  watermark,
  imageSrc,
  context,
}: WatermarkPaintOptions): readonly DisplayPrimitive[] => {
  switch (watermark.kind) {
    case "text":
      return paintTextWatermark(watermark, page, context);
    case "picture":
      return paintPictureWatermark(watermark, page, imageSrc, context);
    default:
      watermark satisfies never;
      return [];
  }
};
