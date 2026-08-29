/**
 * Page-raster comparison for the interoperability harness. Geometry catches
 * text flow defects; this pass catches painted differences such as borders,
 * fills, images, shapes, and glyph rendering.
 */

import { mkdir, readFile, stat } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";
import pixelmatch from "pixelmatch";
import { PNG } from "pngjs";

import type { RasterComparison, RasterPageComparison } from "./types";

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const MAX_RASTER_PAGES = 250;
const MAX_PNG_BYTES = 64 * 1024 * 1024;
const MAX_RASTER_PIXELS = 25_000_000;
const MAX_TOTAL_PNG_BYTES = 512 * 1024 * 1024;
const MAX_TOTAL_RASTER_PIXELS = 250_000_000;
const MAX_DIFF_PNG_BYTES = 64 * 1024 * 1024;
const MAX_TOTAL_DIFF_BYTES = 256 * 1024 * 1024;
const PIXELMATCH_THRESHOLD = 0.1;

/** Signals invalid or resource-exceeding raster comparison input. */
export class RasterComparisonError extends TaggedError("RasterComparisonError")<{
  message: string;
}> {}

const rasterError = (message: string): RasterComparisonError =>
  new RasterComparisonError({ message });

type ComparePageRastersOptions = {
  referencePagePngs: string[];
  folioPagePngs: string[];
  outputDir: string;
};

type DecodedPng = {
  width: number;
  height: number;
  data: Buffer;
  sourceBytes: number;
};

type PngDimensions = { width: number; height: number };

const readPngDimensions = (bytes: Buffer, filePath: string): PngDimensions => {
  if (bytes.length < 24 || !bytes.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)) {
    throw rasterError(`Invalid PNG generated for parity comparison: ${filePath}`);
  }

  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  if (width === 0 || height === 0 || width * height > MAX_RASTER_PIXELS) {
    throw rasterError(
      `Parity PNG dimensions exceed the ${MAX_RASTER_PIXELS.toLocaleString("en")} pixel limit: ${width}x${height} (${filePath})`,
    );
  }
  return { width, height };
};

const decodePng = async (filePath: string): Promise<DecodedPng> => {
  const file = await stat(filePath);
  if (!file.isFile() || file.size > MAX_PNG_BYTES) {
    throw rasterError(
      `Parity PNG exceeds the ${MAX_PNG_BYTES.toLocaleString("en")} byte input limit: ${filePath}`,
    );
  }
  const bytes = await readFile(filePath);
  if (bytes.byteLength > MAX_PNG_BYTES) {
    throw rasterError(
      `Parity PNG exceeds the ${MAX_PNG_BYTES.toLocaleString("en")} byte input limit: ${filePath}`,
    );
  }
  const expected = readPngDimensions(bytes, filePath);
  const decoded = PNG.sync.read(bytes);
  if (decoded.width !== expected.width || decoded.height !== expected.height) {
    throw rasterError(`PNG dimensions changed while decoding: ${filePath}`);
  }
  return { ...decoded, sourceBytes: bytes.byteLength };
};

const writeDiffPng = async (
  diff: PNG,
  filePath: string,
  remainingBytes: number,
): Promise<number> => {
  const bytes = PNG.sync.write(diff);
  if (bytes.byteLength > MAX_DIFF_PNG_BYTES || bytes.byteLength > remainingBytes) {
    throw rasterError(`Raster diff output exceeds its byte budget: ${filePath}`);
  }
  await Bun.write(filePath, bytes);
  return bytes.byteLength;
};

const whiteCanvas = (width: number, height: number): Buffer => {
  const data = Buffer.alloc(width * height * 4, 255);
  return data;
};

const fitToCanvas = (source: DecodedPng, width: number, height: number): Buffer => {
  if (source.width === width && source.height === height) {
    return source.data;
  }

  const padded = whiteCanvas(width, height);
  const sourceStride = source.width * 4;
  const targetStride = width * 4;
  const copiedBytesPerRow = Math.min(source.width, width) * 4;
  for (let row = 0; row < Math.min(source.height, height); row += 1) {
    const sourceStart = row * sourceStride;
    source.data.copy(padded, row * targetStride, sourceStart, sourceStart + copiedBytesPerRow);
  }
  return padded;
};

const markOutsideOverlap = (diff: PNG, overlapWidth: number, overlapHeight: number): void => {
  for (let y = 0; y < diff.height; y += 1) {
    for (let x = 0; x < diff.width; x += 1) {
      if (x < overlapWidth && y < overlapHeight) continue;
      const offset = (y * diff.width + x) * 4;
      diff.data[offset] = 207;
      diff.data[offset + 1] = 34;
      diff.data[offset + 2] = 46;
      diff.data[offset + 3] = 255;
    }
  }
};

const missingPageDiff = (width: number, height: number): PNG => {
  const diff = new PNG({ width, height });
  for (let offset = 0; offset < diff.data.length; offset += 4) {
    diff.data[offset] = 207;
    diff.data[offset + 1] = 34;
    diff.data[offset + 2] = 46;
    diff.data[offset + 3] = 255;
  }
  return diff;
};

type ComparePresentPageOptions = {
  page: number;
  reference: DecodedPng;
  folio: DecodedPng;
  diffPath: string;
  remainingOutputBytes: number;
};

type ComparedPage = {
  comparison: RasterPageComparison;
  outputBytes: number;
};

const comparePresentPage = async ({
  page,
  reference,
  folio,
  diffPath,
  remainingOutputBytes,
}: ComparePresentPageOptions): Promise<ComparedPage> => {
  const width = Math.max(reference.width, folio.width);
  const height = Math.max(reference.height, folio.height);
  const referenceData = fitToCanvas(reference, width, height);
  const folioData = fitToCanvas(folio, width, height);
  const diff = new PNG({ width, height });
  const pixelmatchOptions = {
    threshold: PIXELMATCH_THRESHOLD,
    includeAA: false,
    alpha: 0.7,
    diffColor: [207, 34, 46],
    aaColor: [154, 103, 0],
  } satisfies NonNullable<Parameters<typeof pixelmatch>[5]>;
  const paintedDiffPixels = pixelmatch(
    referenceData,
    folioData,
    diff.data,
    width,
    height,
    pixelmatchOptions,
  );

  const totalPixels = width * height;
  if (reference.width !== folio.width || reference.height !== folio.height) {
    const overlapWidth = Math.min(reference.width, folio.width);
    const overlapHeight = Math.min(reference.height, folio.height);
    const overlapPixels = overlapWidth * overlapHeight;
    const overlapDiffPixels = pixelmatch(
      fitToCanvas(reference, overlapWidth, overlapHeight),
      fitToCanvas(folio, overlapWidth, overlapHeight),
      undefined,
      overlapWidth,
      overlapHeight,
      pixelmatchOptions,
    );
    const diffPixels = overlapDiffPixels + totalPixels - overlapPixels;
    markOutsideOverlap(diff, overlapWidth, overlapHeight);
    const outputBytes = await writeDiffPng(diff, diffPath, remainingOutputBytes);
    return {
      comparison: {
        status: "dimension-mismatch",
        page,
        referenceWidthPx: reference.width,
        referenceHeightPx: reference.height,
        folioWidthPx: folio.width,
        folioHeightPx: folio.height,
        diffPixels,
        totalPixels,
        similarity: 1 - diffPixels / totalPixels,
      },
      outputBytes,
    };
  }
  const outputBytes = await writeDiffPng(diff, diffPath, remainingOutputBytes);
  const diffPixels = paintedDiffPixels;
  const similarity = 1 - diffPixels / totalPixels;
  if (diffPixels === 0) {
    return {
      comparison: {
        status: "match",
        page,
        widthPx: width,
        heightPx: height,
        diffPixels,
        totalPixels,
        similarity,
      },
      outputBytes,
    };
  }
  return {
    comparison: {
      status: "difference",
      page,
      widthPx: width,
      heightPx: height,
      diffPixels,
      totalPixels,
      similarity,
    },
    outputBytes,
  };
};

type CompareMissingPageOptions = {
  page: number;
  present: DecodedPng;
  missing: "reference" | "folio";
  diffPath: string;
  remainingOutputBytes: number;
};

const compareMissingPage = async ({
  page,
  present,
  missing,
  diffPath,
  remainingOutputBytes,
}: CompareMissingPageOptions): Promise<ComparedPage> => {
  const outputBytes = await writeDiffPng(
    missingPageDiff(present.width, present.height),
    diffPath,
    remainingOutputBytes,
  );
  const totalPixels = present.width * present.height;
  return {
    comparison: {
      status: missing === "reference" ? "missing-reference" : "missing-folio",
      page,
      widthPx: present.width,
      heightPx: present.height,
      diffPixels: totalPixels,
      totalPixels,
      similarity: 0,
    },
    outputBytes,
  };
};

export type ComparePageRastersResult = {
  comparison: RasterComparison;
  diffPagePngs: string[];
};

/** Compare corresponding reference/Folio pages and write one diagnostic diff
 * PNG per page. Missing pages count as fully different; different dimensions
 * are padded with white so the size error remains visible in the diff. */
export const comparePageRasters = async ({
  referencePagePngs,
  folioPagePngs,
  outputDir,
}: ComparePageRastersOptions): Promise<ComparePageRastersResult> => {
  const pageCount = Math.max(referencePagePngs.length, folioPagePngs.length);
  if (pageCount > MAX_RASTER_PAGES) {
    throw rasterError(
      `Raster comparison exceeds the ${MAX_RASTER_PAGES.toLocaleString("en")} page limit`,
    );
  }
  if (pageCount === 0) {
    return { comparison: { status: "empty", pages: [] }, diffPagePngs: [] };
  }
  await mkdir(outputDir, { recursive: true });

  const pages: RasterPageComparison[] = [];
  const diffPagePngs: string[] = [];
  let totalInputBytes = 0;
  let totalRasterPixels = 0;
  let totalOutputBytes = 0;
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    const page = pageIndex + 1;
    const referencePath = referencePagePngs[pageIndex];
    const folioPath = folioPagePngs[pageIndex];
    const diffPath = path.join(outputDir, `p${page}.png`);

    let result: ComparedPage;
    if (referencePath !== undefined && folioPath !== undefined) {
      // oxlint-disable-next-line no-await-in-loop -- bound page rasters are decoded and released sequentially
      const [reference, folio] = await Promise.all([
        decodePng(referencePath),
        decodePng(folioPath),
      ]);
      totalInputBytes += reference.sourceBytes + folio.sourceBytes;
      totalRasterPixels +=
        Math.max(reference.width, folio.width) * Math.max(reference.height, folio.height);
      if (totalInputBytes > MAX_TOTAL_PNG_BYTES || totalRasterPixels > MAX_TOTAL_RASTER_PIXELS) {
        throw rasterError("Raster comparison exceeds its aggregate input budget");
      }
      // oxlint-disable-next-line no-await-in-loop -- diff artifacts are emitted in deterministic page order
      result = await comparePresentPage({
        page,
        reference,
        folio,
        diffPath,
        remainingOutputBytes: MAX_TOTAL_DIFF_BYTES - totalOutputBytes,
      });
    } else {
      const presentPath = referencePath ?? folioPath;
      if (presentPath === undefined) {
        throw rasterError(`No raster exists for page ${page}`);
      }
      // oxlint-disable-next-line no-await-in-loop -- bound page rasters are decoded and released sequentially
      const present = await decodePng(presentPath);
      totalInputBytes += present.sourceBytes;
      totalRasterPixels += present.width * present.height;
      if (totalInputBytes > MAX_TOTAL_PNG_BYTES || totalRasterPixels > MAX_TOTAL_RASTER_PIXELS) {
        throw rasterError("Raster comparison exceeds its aggregate input budget");
      }
      // oxlint-disable-next-line no-await-in-loop -- diff artifacts are emitted in deterministic page order
      result = await compareMissingPage({
        page,
        present,
        missing: referencePath === undefined ? "reference" : "folio",
        diffPath,
        remainingOutputBytes: MAX_TOTAL_DIFF_BYTES - totalOutputBytes,
      });
    }

    totalOutputBytes += result.outputBytes;
    pages.push(result.comparison);
    diffPagePngs.push(diffPath);
  }

  const totalPixels = pages.reduce((sum, page) => sum + page.totalPixels, 0);
  const diffPixels = pages.reduce((sum, page) => sum + page.diffPixels, 0);
  return {
    comparison: {
      status: "compared",
      score: 1 - diffPixels / totalPixels,
      diffPixels,
      totalPixels,
      pages,
    },
    diffPagePngs,
  };
};
