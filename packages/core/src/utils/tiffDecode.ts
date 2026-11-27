/**
 * Decoding the first image of a TIFF to RGBA, given the decoder.
 *
 * The decoder (`utif2`) is not imported here: the TIFF worker
 * (`tiffDecoder.worker.ts`) imports it, so it loads only where a TIFF is
 * decoded and no editor bundle carries it, and a host may hand one over
 * instead (`configureTiffDecoding`).
 */

import type * as Utif from "utif2";

/** The part of `utif2` that decoding uses; `import("utif2")` satisfies it. */
export type TiffDecoder = Pick<
  typeof Utif,
  "decode" | "decodeImage" | "toRGBA8"
>;

const hasDecoderFunctions = (value: unknown): value is TiffDecoder =>
  (typeof value === "object" || typeof value === "function") &&
  value !== null &&
  "decode" in value &&
  typeof value.decode === "function" &&
  "decodeImage" in value &&
  typeof value.decodeImage === "function" &&
  "toRGBA8" in value &&
  typeof value.toRGBA8 === "function";

/**
 * The decoder in an imported `utif2` module. It is CommonJS, which bundlers
 * expose either as the namespace itself or as its `default`.
 */
export const tiffDecoderOf = (imported: unknown): TiffDecoder | undefined => {
  if (hasDecoderFunctions(imported)) {
    return imported;
  }
  const fallback: unknown =
    typeof imported === "object" && imported !== null && "default" in imported
      ? imported.default
      : undefined;
  return hasDecoderFunctions(fallback) ? fallback : undefined;
};

export type DecodedTiff = {
  readonly width: number;
  readonly height: number;
  readonly rgba: Uint8Array;
};

/**
 * Read a numeric TIFF tag value. utif2 stores tag values as arrays (typed or
 * plain), so the cap check only needs the first element.
 */
const readTiffTagNumber = (tag: unknown): number | undefined => {
  if (typeof tag === "number") {
    return tag;
  }
  if (
    Array.isArray(tag) ||
    tag instanceof Uint8Array ||
    tag instanceof Uint16Array ||
    tag instanceof Uint32Array
  ) {
    const first = tag[0];
    if (typeof first === "number") {
      return first;
    }
  }
  return undefined;
};

/**
 * The first image of `tiff` as RGBA, or `null` when it has none or declares
 * more than `maxPixels` pixels. Throws what the decoder throws.
 */
export const decodeFirstTiffImage = (
  decoder: TiffDecoder,
  tiff: ArrayBuffer,
  maxPixels: number,
): DecodedTiff | null => {
  const firstImage = decoder.decode(tiff)[0];
  if (!firstImage) {
    return null;
  }

  // `decode()` returns IFDs whose dimensions are raw tags (`t256` =
  // ImageWidth, `t257` = ImageLength); `width` / `height` are set only by
  // `decodeImage()`. Read the tags so the pixel cap holds BEFORE the eager
  // RGBA allocation runs.
  const declaredWidth = readTiffTagNumber(firstImage["t256"]);
  const declaredHeight = readTiffTagNumber(firstImage["t257"]);
  if (!declaredWidth || !declaredHeight) {
    return null;
  }
  if (declaredWidth * declaredHeight > maxPixels) {
    return null;
  }

  decoder.decodeImage(tiff, firstImage);
  const rgba = decoder.toRGBA8(firstImage);
  // The decoded dimensions may differ from the tags after orientation or
  // strip handling.
  const { width, height } = firstImage;
  if (rgba.length === 0 || !width || !height) {
    return null;
  }
  return { width, height, rgba };
};

/** Whether a worker's reply is a decoded image. */
export const isDecodedTiff = (value: unknown): value is DecodedTiff =>
  typeof value === "object" &&
  value !== null &&
  "width" in value &&
  typeof value.width === "number" &&
  "height" in value &&
  typeof value.height === "number" &&
  "rgba" in value &&
  value.rgba instanceof Uint8Array &&
  value.rgba.length === value.width * value.height * 4;
