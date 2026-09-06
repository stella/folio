/**
 * Image decoding for PDF image XObjects.
 *
 * JPEG bytes go into the file untouched under `/DCTDecode`: re-encoding a
 * lossy image to paint it is a fidelity loss with nothing to gain. PNG is
 * decoded only as far as PDF needs, which is less far than a rasterizer
 * would go: the filters are reversed and the samples handed over as they are,
 * so a 4-bit palette image stays 4-bit and an alpha channel becomes an
 * `/SMask` rather than being composited against a guess at the backdrop.
 *
 * A form this module cannot represent exactly (16-bit samples, interlacing)
 * is an error, never an approximation: painting the wrong pixels silently is
 * worse than refusing the page.
 */

import { Result, TaggedError } from "better-result";
import { inflateSync } from "node:zlib";
import type { DisplayImageSource } from "../display-list/types";
import { pdfArray, pdfHexString, pdfName, pdfNumber, type PdfValue } from "./objects";

export class PdfImageError extends TaggedError("PdfImageError")<{ message: string }> {}

/** A soft mask: one alpha sample per pixel, matching the image's dimensions. */
export type PdfImageAlpha = {
  readonly data: Uint8Array;
  readonly bitsPerComponent: number;
};

export type PdfImageSamples = {
  readonly widthPx: number;
  readonly heightPx: number;
  readonly bitsPerComponent: number;
  readonly colorSpace: PdfValue;
  /** `jpeg` is written as `/DCTDecode`; `raw` is deflated by the writer. */
  readonly encoding: "raw" | "jpeg";
  readonly data: Uint8Array;
  /** Component inversion, for Adobe CMYK JPEGs. */
  readonly decode?: readonly number[];
  readonly alpha?: PdfImageAlpha;
};

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a] as const;

const PNG_COLOR_TYPE = {
  greyscale: 0,
  rgb: 2,
  palette: 3,
  greyscaleAlpha: 4,
  rgba: 6,
} as const;

const CHANNELS_PER_COLOR_TYPE = {
  [PNG_COLOR_TYPE.greyscale]: 1,
  [PNG_COLOR_TYPE.rgb]: 3,
  [PNG_COLOR_TYPE.palette]: 1,
  [PNG_COLOR_TYPE.greyscaleAlpha]: 2,
  [PNG_COLOR_TYPE.rgba]: 4,
} as const satisfies Record<number, number>;

const FULLY_OPAQUE = 0xff;
const EIGHT_BIT = 8;

const toHex = (bytes: Uint8Array): string => {
  let out = "";
  for (const byte of bytes) {
    out += byte.toString(16).padStart(2, "0").toUpperCase();
  }
  return out;
};

type PngChunks = {
  readonly header: Uint8Array;
  readonly palette: Uint8Array | null;
  readonly transparency: Uint8Array | null;
  readonly pixels: Uint8Array;
};

const readPngChunks = (bytes: Uint8Array): Result<PngChunks, PdfImageError> => {
  for (let index = 0; index < PNG_SIGNATURE.length; index += 1) {
    if (bytes[index] !== PNG_SIGNATURE[index]) {
      return Result.err(new PdfImageError({ message: "not a PNG: bad signature" }));
    }
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let header: Uint8Array | null = null;
  let palette: Uint8Array | null = null;
  let transparency: Uint8Array | null = null;
  const idat: Uint8Array[] = [];
  let cursor = PNG_SIGNATURE.length;
  const CHUNK_HEADER_BYTES = 8;
  const CHUNK_CRC_BYTES = 4;

  while (cursor + CHUNK_HEADER_BYTES <= bytes.length) {
    const length = view.getUint32(cursor);
    const type = String.fromCharCode(
      bytes[cursor + 4] ?? 0,
      bytes[cursor + 5] ?? 0,
      bytes[cursor + 6] ?? 0,
      bytes[cursor + 7] ?? 0,
    );
    const dataStart = cursor + CHUNK_HEADER_BYTES;
    const dataEnd = dataStart + length;
    if (dataEnd > bytes.length) {
      return Result.err(new PdfImageError({ message: `PNG chunk ${type} runs past the end` }));
    }
    const data = bytes.subarray(dataStart, dataEnd);
    if (type === "IHDR") {
      header = data;
    } else if (type === "PLTE") {
      palette = data;
    } else if (type === "tRNS") {
      transparency = data;
    } else if (type === "IDAT") {
      idat.push(data);
    } else if (type === "IEND") {
      break;
    }
    cursor = dataEnd + CHUNK_CRC_BYTES;
  }

  if (header === null || header.length < 13) {
    return Result.err(new PdfImageError({ message: "PNG has no IHDR" }));
  }
  if (idat.length === 0) {
    return Result.err(new PdfImageError({ message: "PNG has no IDAT" }));
  }
  const totalPixelBytes = idat.reduce((sum, chunk) => sum + chunk.length, 0);
  const pixels = new Uint8Array(totalPixelBytes);
  let pixelCursor = 0;
  for (const chunk of idat) {
    pixels.set(chunk, pixelCursor);
    pixelCursor += chunk.length;
  }
  return Result.ok({ header, palette, transparency, pixels });
};

const paeth = (left: number, above: number, upperLeft: number): number => {
  const estimate = left + above - upperLeft;
  const distanceLeft = Math.abs(estimate - left);
  const distanceAbove = Math.abs(estimate - above);
  const distanceUpperLeft = Math.abs(estimate - upperLeft);
  if (distanceLeft <= distanceAbove && distanceLeft <= distanceUpperLeft) {
    return left;
  }
  return distanceAbove <= distanceUpperLeft ? above : upperLeft;
};

type UnfilterOptions = {
  readonly filtered: Uint8Array;
  readonly bytesPerRow: number;
  readonly bytesPerPixel: number;
  readonly height: number;
};

/**
 * Reverses the five PNG filter types in place over a copy, leaving the
 * scanlines byte-aligned exactly as PDF expects them.
 */
const unfilterPng = ({
  filtered,
  bytesPerRow,
  bytesPerPixel,
  height,
}: UnfilterOptions): Result<Uint8Array, PdfImageError> => {
  const expected = height * (bytesPerRow + 1);
  if (filtered.length < expected) {
    return Result.err(
      new PdfImageError({
        message: `PNG pixel data is ${String(filtered.length)} bytes, expected ${String(expected)}`,
      }),
    );
  }
  const out = new Uint8Array(height * bytesPerRow);
  for (let row = 0; row < height; row += 1) {
    const filterType = filtered[row * (bytesPerRow + 1)] ?? 0;
    const source = row * (bytesPerRow + 1) + 1;
    const target = row * bytesPerRow;
    const previous = target - bytesPerRow;
    for (let index = 0; index < bytesPerRow; index += 1) {
      const raw = filtered[source + index] ?? 0;
      const left = index >= bytesPerPixel ? (out[target + index - bytesPerPixel] ?? 0) : 0;
      const above = row > 0 ? (out[previous + index] ?? 0) : 0;
      const upperLeft =
        row > 0 && index >= bytesPerPixel ? (out[previous + index - bytesPerPixel] ?? 0) : 0;
      let value: number;
      switch (filterType) {
        case 0:
          value = raw;
          break;
        case 1:
          value = raw + left;
          break;
        case 2:
          value = raw + above;
          break;
        case 3:
          value = raw + ((left + above) >> 1);
          break;
        case 4:
          value = raw + paeth(left, above, upperLeft);
          break;
        default:
          return Result.err(
            new PdfImageError({ message: `unknown PNG filter type ${String(filterType)}` }),
          );
      }
      out[target + index] = value & 0xff;
    }
  }
  return Result.ok(out);
};

type ExpandIndicesOptions = {
  readonly rows: Uint8Array;
  readonly bytesPerRow: number;
  readonly width: number;
  readonly height: number;
  readonly bitDepth: number;
};

/** One byte per pixel index, for sub-byte palette and greyscale depths. */
const expandIndices = ({
  rows,
  bytesPerRow,
  width,
  height,
  bitDepth,
}: ExpandIndicesOptions): Uint8Array => {
  if (bitDepth === EIGHT_BIT) {
    return rows;
  }
  const out = new Uint8Array(width * height);
  const perByte = EIGHT_BIT / bitDepth;
  const mask = (1 << bitDepth) - 1;
  for (let row = 0; row < height; row += 1) {
    for (let column = 0; column < width; column += 1) {
      const byte = rows[row * bytesPerRow + Math.floor(column / perByte)] ?? 0;
      const shift = EIGHT_BIT - bitDepth * ((column % perByte) + 1);
      out[row * width + column] = (byte >> shift) & mask;
    }
  }
  return out;
};

const decodePng = (bytes: Uint8Array): Result<PdfImageSamples, PdfImageError> => {
  const chunks = readPngChunks(bytes);
  if (chunks.isErr()) {
    return Result.err(chunks.error);
  }
  const { header, palette, transparency, pixels } = chunks.value;
  const headerView = new DataView(header.buffer, header.byteOffset, header.byteLength);
  const width = headerView.getUint32(0);
  const height = headerView.getUint32(4);
  const bitDepth = header[8] ?? 0;
  const colorType = header[9] ?? 0;
  const interlace = header[12] ?? 0;

  if (width === 0 || height === 0) {
    return Result.err(new PdfImageError({ message: "PNG has a zero dimension" }));
  }
  if (interlace !== 0) {
    return Result.err(new PdfImageError({ message: "interlaced PNG is not supported" }));
  }
  const channels = CHANNELS_PER_COLOR_TYPE[colorType as keyof typeof CHANNELS_PER_COLOR_TYPE];
  if (channels === undefined) {
    return Result.err(
      new PdfImageError({ message: `unknown PNG colour type ${String(colorType)}` }),
    );
  }
  const packedDepthAllowed =
    colorType === PNG_COLOR_TYPE.greyscale || colorType === PNG_COLOR_TYPE.palette;
  const depthAllowed = packedDepthAllowed
    ? [1, 2, 4, EIGHT_BIT].includes(bitDepth)
    : bitDepth === EIGHT_BIT;
  if (!depthAllowed) {
    return Result.err(
      new PdfImageError({
        message: `PNG bit depth ${String(bitDepth)} with colour type ${String(colorType)} is not supported`,
      }),
    );
  }

  const inflated = inflateSync(pixels);
  const filtered = new Uint8Array(inflated.buffer, inflated.byteOffset, inflated.byteLength);
  const bitsPerPixel = bitDepth * channels;
  const bytesPerRow = Math.ceil((bitsPerPixel * width) / EIGHT_BIT);
  const bytesPerPixel = Math.max(1, Math.ceil(bitsPerPixel / EIGHT_BIT));
  const unfiltered = unfilterPng({ filtered, bytesPerRow, bytesPerPixel, height });
  if (unfiltered.isErr()) {
    return Result.err(unfiltered.error);
  }
  const rows = unfiltered.value;
  const pixelCount = width * height;

  switch (colorType) {
    case PNG_COLOR_TYPE.greyscale:
    case PNG_COLOR_TYPE.rgb: {
      const isRgb = colorType === PNG_COLOR_TYPE.rgb;
      let alpha: PdfImageAlpha | undefined;
      if (transparency !== null) {
        const built = buildColorKeyAlpha({
          rows,
          bitDepth,
          bytesPerRow,
          width,
          height,
          components: isRgb ? 3 : 1,
          key: isRgb
            ? [transparency[1] ?? 0, transparency[3] ?? 0, transparency[5] ?? 0]
            : [transparency[1] ?? 0],
        });
        if (built.isErr()) {
          return Result.err(built.error);
        }
        alpha = built.value;
      }
      return Result.ok({
        widthPx: width,
        heightPx: height,
        bitsPerComponent: isRgb ? EIGHT_BIT : bitDepth,
        colorSpace: pdfName(isRgb ? "DeviceRGB" : "DeviceGray"),
        encoding: "raw",
        data: rows,
        ...(alpha === undefined ? {} : { alpha }),
      });
    }
    case PNG_COLOR_TYPE.palette: {
      if (palette === null) {
        return Result.err(new PdfImageError({ message: "palette PNG has no PLTE chunk" }));
      }
      const entries = Math.floor(palette.length / 3);
      const colorSpace = pdfArray([
        pdfName("Indexed"),
        pdfName("DeviceRGB"),
        pdfNumber(entries - 1),
        pdfHexString(toHex(palette)),
      ]);
      let alpha: PdfImageAlpha | undefined;
      if (transparency !== null) {
        const indices = expandIndices({ rows, bytesPerRow, width, height, bitDepth });
        const data = new Uint8Array(pixelCount);
        for (let index = 0; index < pixelCount; index += 1) {
          const paletteIndex = indices[index] ?? 0;
          data[index] = transparency[paletteIndex] ?? FULLY_OPAQUE;
        }
        alpha = { data, bitsPerComponent: EIGHT_BIT };
      }
      return Result.ok({
        widthPx: width,
        heightPx: height,
        bitsPerComponent: bitDepth,
        colorSpace,
        encoding: "raw",
        data: rows,
        ...(alpha === undefined ? {} : { alpha }),
      });
    }
    case PNG_COLOR_TYPE.greyscaleAlpha:
    case PNG_COLOR_TYPE.rgba: {
      const colorComponents = colorType === PNG_COLOR_TYPE.rgba ? 3 : 1;
      const color = new Uint8Array(pixelCount * colorComponents);
      const alphaData = new Uint8Array(pixelCount);
      for (let index = 0; index < pixelCount; index += 1) {
        const source = index * (colorComponents + 1);
        for (let component = 0; component < colorComponents; component += 1) {
          color[index * colorComponents + component] = rows[source + component] ?? 0;
        }
        alphaData[index] = rows[source + colorComponents] ?? FULLY_OPAQUE;
      }
      return Result.ok({
        widthPx: width,
        heightPx: height,
        bitsPerComponent: EIGHT_BIT,
        colorSpace: pdfName(colorComponents === 3 ? "DeviceRGB" : "DeviceGray"),
        encoding: "raw",
        data: color,
        alpha: { data: alphaData, bitsPerComponent: EIGHT_BIT },
      });
    }
    default:
      return Result.err(
        new PdfImageError({ message: `unknown PNG colour type ${String(colorType)}` }),
      );
  }
};

type ColorKeyAlphaOptions = {
  readonly rows: Uint8Array;
  readonly bitDepth: number;
  readonly bytesPerRow: number;
  readonly width: number;
  readonly height: number;
  readonly components: number;
  readonly key: readonly number[];
};

/**
 * `tRNS` on a truecolour or greyscale PNG names one fully transparent sample
 * value. PDF has no colour-key transparency for image XObjects, so it becomes
 * a soft mask.
 */
const buildColorKeyAlpha = ({
  rows,
  bitDepth,
  bytesPerRow,
  width,
  height,
  components,
  key,
}: ColorKeyAlphaOptions): Result<PdfImageAlpha, PdfImageError> => {
  if (bitDepth !== EIGHT_BIT) {
    return Result.err(
      new PdfImageError({
        message: `tRNS on a ${String(bitDepth)}-bit non-palette PNG is not supported`,
      }),
    );
  }
  const data = new Uint8Array(width * height);
  for (let row = 0; row < height; row += 1) {
    for (let column = 0; column < width; column += 1) {
      const source = row * bytesPerRow + column * components;
      let matches = true;
      for (let component = 0; component < components; component += 1) {
        if ((rows[source + component] ?? 0) !== (key[component] ?? 0)) {
          matches = false;
          break;
        }
      }
      data[row * width + column] = matches ? 0 : FULLY_OPAQUE;
    }
  }
  return Result.ok({ data, bitsPerComponent: EIGHT_BIT });
};

const JPEG_COMPONENTS_TO_COLOR_SPACE = {
  1: "DeviceGray",
  3: "DeviceRGB",
  4: "DeviceCMYK",
} as const satisfies Record<number, string>;

const CMYK_INVERTED_DECODE = [1, 0, 1, 0, 1, 0, 1, 0] as const;

const decodeJpeg = (bytes: Uint8Array): Result<PdfImageSamples, PdfImageError> => {
  const MARKER_PREFIX = 0xff;
  const START_OF_IMAGE = 0xd8;
  const START_OF_SCAN = 0xda;
  const DEFINE_HUFFMAN_TABLES = 0xc4;
  const DEFINE_ARITHMETIC_CONDITIONING = 0xcc;
  const APP14 = 0xee;

  if (bytes[0] !== MARKER_PREFIX || bytes[1] !== START_OF_IMAGE) {
    return Result.err(new PdfImageError({ message: "not a JPEG: bad SOI marker" }));
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let cursor = 2;
  let adobe = false;
  while (cursor + 4 <= bytes.length) {
    if (bytes[cursor] !== MARKER_PREFIX) {
      return Result.err(new PdfImageError({ message: "JPEG marker segment is out of step" }));
    }
    const marker = bytes[cursor + 1] ?? 0;
    if (marker === START_OF_SCAN) {
      break;
    }
    const length = view.getUint16(cursor + 2);
    const segment = cursor + 4;
    const isFrameHeader =
      marker >= 0xc0 &&
      marker <= 0xcf &&
      marker !== DEFINE_HUFFMAN_TABLES &&
      marker !== DEFINE_ARITHMETIC_CONDITIONING;
    if (marker === APP14 && length >= 12) {
      adobe = String.fromCharCode(...bytes.subarray(segment, segment + 5)) === "Adobe";
    }
    if (isFrameHeader) {
      const height = view.getUint16(segment + 1);
      const width = view.getUint16(segment + 3);
      const components = bytes[segment + 5] ?? 0;
      const colorSpace =
        JPEG_COMPONENTS_TO_COLOR_SPACE[components as keyof typeof JPEG_COMPONENTS_TO_COLOR_SPACE];
      if (colorSpace === undefined) {
        return Result.err(
          new PdfImageError({
            message: `JPEG with ${String(components)} components is not supported`,
          }),
        );
      }
      // Adobe writes CMYK JPEGs with inverted samples; every other producer's
      // CMYK is already the right way round, so the marker decides.
      const inverted = components === 4 && adobe;
      return Result.ok({
        widthPx: width,
        heightPx: height,
        bitsPerComponent: EIGHT_BIT,
        colorSpace: pdfName(colorSpace),
        encoding: "jpeg",
        data: bytes,
        ...(inverted ? { decode: CMYK_INVERTED_DECODE } : {}),
      });
    }
    cursor += 2 + length;
  }
  return Result.err(new PdfImageError({ message: "JPEG has no frame header" }));
};

export const decodeImage = (source: DisplayImageSource): Result<PdfImageSamples, PdfImageError> => {
  switch (source.format) {
    case "png":
      return decodePng(source.bytes);
    case "jpeg":
      return decodeJpeg(source.bytes);
    default: {
      const unreachable: never = source.format;
      return Result.err(
        new PdfImageError({ message: `unsupported image format ${String(unreachable)}` }),
      );
    }
  }
};
