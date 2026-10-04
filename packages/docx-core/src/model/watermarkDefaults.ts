import type { Watermark } from "./content";

export const WATERMARK_DEFAULTS = {
  textFont: "Calibri",
  textColor: "C0C0C0",
  pictureWidthPt: 415,
  pictureHeightPt: 207,
  pictureScaleEpsilon: 0.01,
  pictureWashoutGain: "19661f",
  pictureWashoutBlacklevel: "22938f",
} as const;

/** The canonical model records the properties that watermark synthesis emits. */
export const normalizeCanonicalWatermark = (watermark: Watermark): Watermark => {
  if (watermark.kind === "text")
    return {
      kind: "text",
      text: watermark.text,
      font: watermark.font ?? WATERMARK_DEFAULTS.textFont,
      color:
        watermark.color && watermark.color !== "auto"
          ? watermark.color.toUpperCase()
          : WATERMARK_DEFAULTS.textColor,
      diagonal: watermark.diagonal ?? true,
      ...(watermark.opacity === undefined ? {} : { opacity: watermark.opacity }),
    };
  const widthPt =
    watermark.scale === undefined
      ? (watermark.widthPt ?? WATERMARK_DEFAULTS.pictureWidthPt)
      : watermark.scale * WATERMARK_DEFAULTS.pictureWidthPt;
  const heightPt =
    watermark.scale === undefined
      ? (watermark.heightPt ?? WATERMARK_DEFAULTS.pictureHeightPt)
      : watermark.scale * WATERMARK_DEFAULTS.pictureHeightPt;
  const widthScale = widthPt / WATERMARK_DEFAULTS.pictureWidthPt;
  const heightScale = heightPt / WATERMARK_DEFAULTS.pictureHeightPt;
  const uniformScale =
    Math.abs(widthScale - heightScale) > WATERMARK_DEFAULTS.pictureScaleEpsilon
      ? undefined
      : widthScale;
  return {
    kind: "picture",
    imageRId: watermark.imageRId,
    ...(watermark.imageTarget === undefined ? {} : { imageTarget: watermark.imageTarget }),
    ...(watermark.imageTargetExternal === true ? { imageTargetExternal: true } : {}),
    widthPt:
      uniformScale === undefined ? widthPt : uniformScale * WATERMARK_DEFAULTS.pictureWidthPt,
    heightPt:
      uniformScale === undefined ? heightPt : uniformScale * WATERMARK_DEFAULTS.pictureHeightPt,
    ...(uniformScale === undefined ? {} : { scale: uniformScale }),
    ...(watermark.washout === false ? { washout: false } : {}),
  };
};
