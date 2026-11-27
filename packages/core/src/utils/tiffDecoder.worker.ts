/**
 * Worker that decodes a TIFF's first image to RGBA, spawned from
 * `tiffConverter.ts` through `new Worker(new URL(...))`. The decoder loads
 * with this worker, so it is fetched only when a document holds a TIFF.
 *
 * Request: `{ tiff: ArrayBuffer, maxPixels: number }`. Reply: a
 * `DecodedTiff`, its pixels transferred, or `null`.
 */

import * as utif from "utif2";

import {
  type DecodedTiff,
  decodeFirstTiffImage,
  tiffDecoderOf,
} from "./tiffDecode";

const decoder = tiffDecoderOf(utif);

type TiffDecodeRequest = {
  readonly tiff: ArrayBuffer;
  readonly maxPixels: number;
};

const decodeRequest = (request: TiffDecodeRequest): DecodedTiff | null => {
  if (decoder === undefined) {
    return null;
  }
  try {
    return decodeFirstTiffImage(decoder, request.tiff, request.maxPixels);
  } catch {
    return null;
  }
};

// Guarded so this file is importable outside a worker (type checks, tests).
if (typeof self !== "undefined" && typeof addEventListener === "function") {
  // SAFETY: `postMessage` exists in worker scope; this file is only wired up
  // via `new Worker(...)` from `tiffConverter.ts`.
  const workerScope = self as unknown as {
    postMessage: (
      message: DecodedTiff | null,
      transfer: Transferable[],
    ) => void;
  };
  addEventListener("message", (event: MessageEvent<TiffDecodeRequest>) => {
    const decoded = decodeRequest(event.data);
    // eslint-disable-next-line unicorn/require-post-message-target-origin -- worker-scope postMessage takes no targetOrigin.
    workerScope.postMessage(
      decoded,
      decoded === null ? [] : [decoded.rgba.buffer],
    );
  });
}
