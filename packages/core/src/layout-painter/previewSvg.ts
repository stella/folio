/**
 * What the DOM painter shows for a drawing with no image data.
 *
 * Its marks come from the display list's own `paintPreview`, so the two painters
 * cannot draw a preview differently. An SVG in an `<img>` runs no script and
 * fetches nothing, and every value written here is a number.
 */

import { panic } from "better-result";

import { paintPreview } from "../display-list/build/previewPrimitives";
import type { DisplayColor, DisplayPrimitive } from "../display-list/types";
import type { PreviewDescriptor } from "../types/document";

const cssColor = ({ r, g, b, a }: DisplayColor): string =>
  `rgba(${String(r)},${String(g)},${String(b)},${String(a)})`;

const svgPoints = (
  points: readonly { readonly xPx: number; readonly yPx: number }[],
): string =>
  points.map(({ xPx, yPx }) => `${String(xPx)},${String(yPx)}`).join(" ");

const svgMark = (primitive: DisplayPrimitive): string => {
  switch (primitive.kind) {
    case "rect": {
      const { xPx, yPx, widthPx, heightPx } = primitive.rect;
      const fill = primitive.fill ? cssColor(primitive.fill) : "none";
      return `<rect x="${String(xPx)}" y="${String(yPx)}" width="${String(widthPx)}" height="${String(heightPx)}" fill="${fill}"/>`;
    }
    case "polygon":
      return `<polygon points="${svgPoints(primitive.points)}" fill="${cssColor(primitive.fill)}"/>`;
    case "polyline":
      return `<polyline points="${svgPoints(primitive.points)}" fill="none" stroke="${cssColor(primitive.stroke.color)}" stroke-width="${String(primitive.stroke.thicknessPx)}"/>`;
    default:
      return panic(`a preview does not draw ${primitive.kind}`);
  }
};

const dataUrls = new WeakMap<PreviewDescriptor, string>();

/** `preview` as an `<img>` source that fills whatever box the image has. */
export const previewToSvgDataUrl = (preview: PreviewDescriptor): string => {
  const known = dataUrls.get(preview);
  if (known !== undefined) {
    return known;
  }
  const hasArea = preview.extent.width > 0 && preview.extent.height > 0;
  const widthPx = hasArea ? preview.extent.width : 1;
  const heightPx = hasArea ? preview.extent.height : 1;
  const marks = paintPreview(preview, {
    xPx: 0,
    yPx: 0,
    widthPx,
    heightPx,
  });
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${String(widthPx)} ${String(heightPx)}" width="${String(widthPx)}" height="${String(heightPx)}" preserveAspectRatio="none">${marks.map(svgMark).join("")}</svg>`;
  const url = `data:image/svg+xml,${encodeURIComponent(svg)}`;
  dataUrls.set(preview, url);
  return url;
};

/** The source an image paints from: its bytes, or its drawn preview. */
export const paintedImageSource = (image: {
  readonly src: string;
  readonly preview?: PreviewDescriptor;
}): string => (image.preview ? previewToSvgDataUrl(image.preview) : image.src);
