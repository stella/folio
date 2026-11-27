import { afterEach, describe, expect, test } from "bun:test";
import * as UTIF from "utif2";

import {
  configureTiffDecoding,
  convertTiffToPngDataUrl,
  TIFF_PLACEHOLDER_DATA_URL,
} from "./tiffConverter";
import {
  decodeFirstTiffImage,
  isDecodedTiff,
  tiffDecoderOf,
} from "./tiffDecode";

// 2x1: one red pixel, one blue pixel, opaque.
const PIXELS = [255, 0, 0, 255, 0, 0, 255, 255];
const twoPixelTiff = (): ArrayBuffer =>
  UTIF.encodeImage(new Uint8Array(PIXELS), 2, 1);

describe("decodeFirstTiffImage", () => {
  test("decodes the first image to RGBA", () => {
    const decoded = decodeFirstTiffImage(UTIF, twoPixelTiff(), 2);

    expect(decoded).not.toBeNull();
    expect(decoded?.width).toBe(2);
    expect(decoded?.height).toBe(1);
    expect(Array.from(decoded?.rgba ?? [])).toEqual(PIXELS);
    expect(isDecodedTiff(decoded)).toBe(true);
  });

  test("refuses an image over the pixel cap before decoding it", () => {
    let decodedImages = 0;
    const decoder = {
      ...UTIF,
      decodeImage: (...args: Parameters<typeof UTIF.decodeImage>) => {
        decodedImages += 1;
        UTIF.decodeImage(...args);
      },
    };

    expect(decodeFirstTiffImage(decoder, twoPixelTiff(), 1)).toBeNull();
    expect(decodedImages).toBe(0);
  });

  test("finds the decoder on the namespace or on its default", () => {
    expect(tiffDecoderOf(UTIF)).toBe(UTIF);
    expect(tiffDecoderOf({ default: UTIF })).toBe(UTIF);
    expect(tiffDecoderOf({ default: {} })).toBeUndefined();
  });

  test("a reply that does not describe its pixels is not a decoded image", () => {
    expect(isDecodedTiff(null)).toBe(false);
    expect(
      isDecodedTiff({ width: 2, height: 1, rgba: new Uint8Array(4) }),
    ).toBe(false);
  });
});

describe("convertTiffToPngDataUrl", () => {
  afterEach(() => {
    configureTiffDecoding({});
  });

  test("converts nothing without a canvas to paint into", async () => {
    configureTiffDecoding({ decoder: UTIF });

    expect(await convertTiffToPngDataUrl(twoPixelTiff())).toBeNull();
  });

  test("a decoder that fails settles as no conversion", async () => {
    const realDocument: unknown = Reflect.get(globalThis, "document");
    Reflect.set(globalThis, "document", { createElement: () => ({}) });
    try {
      configureTiffDecoding({
        decoder: {
          ...UTIF,
          decode: () => {
            throw new Error("not a TIFF");
          },
        },
      });

      expect(await convertTiffToPngDataUrl(twoPixelTiff())).toBeNull();
    } finally {
      Reflect.set(globalThis, "document", realDocument);
    }
  });

  test("the placeholder is an image source a painter accepts", () => {
    expect(TIFF_PLACEHOLDER_DATA_URL.startsWith("data:image/svg+xml,")).toBe(
      true,
    );
  });
});
