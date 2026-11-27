import { base64ToBytes } from "./textEncoding";
import {
  type DecodedTiff,
  decodeFirstTiffImage,
  isDecodedTiff,
  type TiffDecoder,
  tiffDecoderOf,
} from "./tiffDecode";

/**
 * TIFF to PNG Converter
 *
 * Browsers don't support TIFF in `<img>` tags, so DOCX-embedded TIFFs are
 * decoded and re-encoded as PNG via the Canvas API. The decoder runs in a
 * worker (`tiffDecoder.worker.ts`) fetched the first time a document holds a
 * TIFF, so no editor bundle carries it; a host that cannot load the worker
 * (a single-file bundle, a Content Security Policy) hands the decoder over
 * through {@link configureTiffDecoding}. Returns null in environments without
 * Canvas (Node.js parsing) and when decoding fails; callers keep the TIFF.
 */

/**
 * Hard cap on the pixel count we'll attempt to decode. Each pixel costs 4 B
 * for the intermediate RGBA buffer plus another canvas-managed bitmap, so
 * 64 MP ≈ 256 MB before counting the PNG re-encode. A crafted TIFF whose
 * declared dimensions exceed this is rejected without allocating the
 * RGBA buffer, which would otherwise hang or OOM the tab.
 */
const MAX_TIFF_PIXELS = 64_000_000;

/**
 * Budget shared by every TIFF in one package. The per-image cap bounds a single
 * decode; without a running total a package of many large TIFFs still adds up.
 * Callers pass what is left of the budget and subtract the pixels each
 * conversion reports.
 */
export const MAX_PACKAGE_TIFF_PIXELS = 256_000_000;

/**
 * How long a worker may take to decode one image before the conversion gives
 * up, so a worker that never answers cannot hold a document's opening.
 */
const TIFF_WORKER_TIMEOUT_MS = 30_000;

export type TiffDecodingOptions = {
  /**
   * Decode on the calling thread with this decoder instead of in the
   * packaged worker: the module `await import("utif2")` gives, for hosts
   * that cannot load the worker (offline, a Content Security Policy without
   * `worker-src`, a single-file bundle).
   */
  readonly decoder?: TiffDecoder | { readonly default: TiffDecoder };
};

let decodingOptions: TiffDecodingOptions = {};

/** Set how TIFF images are decoded; each call replaces the previous options. */
export const configureTiffDecoding = (options: TiffDecodingOptions): void => {
  decodingOptions = options;
};

/**
 * What a TIFF paints as when it cannot be decoded: a plain box filling the
 * image's frame, in place of the browser's broken-image icon. The package
 * keeps the TIFF itself.
 */
export const TIFF_PLACEHOLDER_DATA_URL = `data:image/svg+xml,${encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none"><rect width="100" height="100" fill="#f1f3f4" stroke="#c4c7c5" stroke-width="2" vector-effect="non-scaling-stroke"/></svg>',
)}`;

export function isTiffMimeType(mimeType: string): boolean {
  const lower = mimeType.toLowerCase();
  return lower === "image/tiff" || lower === "image/tif";
}

/** Whether this environment can paint a converted TIFF at all. */
export const canConvertTiff = (): boolean =>
  typeof document !== "undefined" &&
  typeof document.createElement === "function";

export type ConvertedTiff = {
  dataUrl: string;
  data: ArrayBuffer;
  /** Pixels decoded, so callers can draw down a package-wide budget. */
  pixels: number;
};

const decodeInWorker = (
  tiff: ArrayBuffer,
  maxPixels: number,
): Promise<DecodedTiff | null> => {
  if (typeof Worker === "undefined") {
    return Promise.resolve(null);
  }
  return new Promise((resolve) => {
    let worker: Worker;
    try {
      // Bundlers resolve the worker URL at build time and emit the worker
      // with its decoder as a chunk of its own.
      worker = new Worker(new URL("tiffDecoder.worker.js", import.meta.url), {
        type: "module",
      });
    } catch {
      resolve(null);
      return;
    }
    const settle = (decoded: DecodedTiff | null): void => {
      clearTimeout(timer);
      worker.terminate();
      resolve(decoded);
    };
    const timer = setTimeout(() => settle(null), TIFF_WORKER_TIMEOUT_MS);
    worker.addEventListener("message", (event: MessageEvent<unknown>) => {
      settle(isDecodedTiff(event.data) ? event.data : null);
    });
    worker.addEventListener("error", () => settle(null));
    worker.addEventListener("messageerror", () => settle(null));
    // Copied, not transferred: the package keeps its bytes.
    // eslint-disable-next-line unicorn/require-post-message-target-origin -- Worker.postMessage takes no targetOrigin.
    worker.postMessage({ tiff, maxPixels });
  });
};

const decodeOnThisThread = (
  supplied: NonNullable<TiffDecodingOptions["decoder"]>,
  tiff: ArrayBuffer,
  maxPixels: number,
): DecodedTiff | null => {
  const decoder = tiffDecoderOf(supplied);
  if (decoder === undefined) {
    return null;
  }
  try {
    return decodeFirstTiffImage(decoder, tiff, maxPixels);
  } catch {
    return null;
  }
};

const decodeTiff = (
  tiff: ArrayBuffer,
  maxPixels: number,
): Promise<DecodedTiff | null> => {
  const supplied = decodingOptions.decoder;
  return supplied === undefined
    ? decodeInWorker(tiff, maxPixels)
    : Promise.resolve(decodeOnThisThread(supplied, tiff, maxPixels));
};

export async function convertTiffToPngDataUrl(
  tiffData: ArrayBuffer,
  maxPixels: number = MAX_TIFF_PIXELS,
): Promise<ConvertedTiff | null> {
  // Bail out before any decode if Canvas isn't available — without it we
  // can't produce a PNG anyway, and TIFF decoding + RGBA conversion would
  // allocate hundreds of MB and burn CPU for nothing in headless parses.
  if (!canConvertTiff()) {
    return null;
  }

  try {
    const decoded = await decodeTiff(
      tiffData,
      Math.min(maxPixels, MAX_TIFF_PIXELS),
    );
    if (decoded === null) {
      return null;
    }
    const { width, height, rgba } = decoded;

    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      return null;
    }

    const clamped = new Uint8ClampedArray(rgba.length);
    clamped.set(rgba);
    const imageData = new ImageData(clamped, width, height);
    ctx.putImageData(imageData, 0, 0);

    const dataUrl = canvas.toDataURL("image/png");
    const data = dataUrlToArrayBuffer(dataUrl);
    if (!data) {
      return null;
    }
    return { dataUrl, data, pixels: width * height };
  } catch {
    return null;
  }
}

function dataUrlToArrayBuffer(dataUrl: string): ArrayBuffer | null {
  const base64 = dataUrl.split(",", 2)[1];
  if (!base64) {
    return null;
  }
  try {
    return base64ToBytes(base64).buffer;
  } catch {
    return null;
  }
}
