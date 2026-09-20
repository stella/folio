/**
 * Turn a {@link PreviewDescriptor} into the PNG a backend paints.
 *
 * Split from the diagram reader so the display list can rasterise without
 * importing the OOXML parser, and so the raster is reached only by something
 * that is about to paint. Parsing a package no longer runs any of this: the
 * megapixel buffer, its two checksums and its base64 were most of what a
 * SmartArt package cost, and none of it was owed until a renderer asked.
 */

import type { PreviewDescriptor, PreviewShape } from "../types/document";

export const MAX_PREVIEW_SHAPES = 128;
export const MAX_PREVIEW_PIXELS = 1_440_000;
const MAX_PREVIEW_PAINT_PIXELS = MAX_PREVIEW_PIXELS * 4;

/**
 * The raster dimensions an extent scales to, computed without rasterising.
 *
 * The parser records these on the descriptor so a consumer can size, budget or
 * lay out a preview it never paints.
 */
export const previewRasterSize = (
  width: number,
  height: number,
): { pixelWidth: number; pixelHeight: number } => {
  const scale = Math.min(1, Math.sqrt(MAX_PREVIEW_PIXELS / (width * height)));
  return {
    pixelWidth: Math.max(1, Math.round(width * scale)),
    pixelHeight: Math.max(1, Math.round(height * scale)),
  };
};

/**
 * The preview is a megapixel raster, so both checksums run over megabytes.
 * `for (const byte of bytes)` drives the array iterator protocol once per
 * byte, which profiles as the dominant cost of painting a SmartArt document;
 * indexed loops and a table-driven CRC produce the same numbers without it.
 */
const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value >>> 1) ^ (value & 1 ? 0xedb88320 : 0);
    }
    table[index] = value >>> 0;
  }
  return table;
})();

const crc32 = (bytes: Uint8Array): number => {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    // SAFETY: `index` is below `bytes.length`, and the table covers every byte.
    crc = (crc >>> 8) ^ (CRC32_TABLE[(crc ^ (bytes[index] as number)) & 0xff] as number);
  }
  return (crc ^ 0xffffffff) >>> 0;
};

/**
 * `a` and `b` stay below 2^31 for 5552 iterations from any legal state, so the
 * modulo runs per block rather than per byte.
 */
const ADLER32_BLOCK = 5552;

const adler32 = (bytes: Uint8Array): number => {
  let a = 1;
  let b = 0;
  let index = 0;
  while (index < bytes.length) {
    const end = Math.min(index + ADLER32_BLOCK, bytes.length);
    for (; index < end; index += 1) {
      // SAFETY: `index` is below `end`, itself at most `bytes.length`.
      a += bytes[index] as number;
      b += a;
    }
    a %= 65521;
    b %= 65521;
  }
  return (b << 16) | a;
};

const pngChunk = (type: string, data: Uint8Array): Uint8Array => {
  const name = new TextEncoder().encode(type);
  const output = new Uint8Array(12 + data.length);
  const view = new DataView(output.buffer);
  view.setUint32(0, data.length);
  output.set(name, 4);
  output.set(data, 8);
  view.setUint32(8 + data.length, crc32(output.subarray(4, 8 + data.length)));
  return output;
};

const previewPng = (
  width: number,
  height: number,
  shapes: readonly PreviewShape[],
  w: number,
  h: number,
): Uint8Array => {
  const pixels = new Uint8Array((w * 4 + 1) * h);
  for (let y = 0; y < h; y += 1) {
    const row = y * (w * 4 + 1);
    pixels[row] = 0;
    for (let x = 0; x < w; x += 1) {
      const index = row + 1 + x * 4;
      pixels[index] = 238;
      pixels[index + 1] = 242;
      pixels[index + 2] = 247;
      pixels[index + 3] = 255;
    }
  }
  let paintedPixels = 0;
  for (const shape of shapes) {
    const sx = Math.max(0, Math.round((shape.x / width) * w));
    const sy = Math.max(0, Math.round((shape.y / height) * h));
    const ex = Math.min(w, Math.round(((shape.x + shape.width) / width) * w));
    const ey = Math.min(h, Math.round(((shape.y + shape.height) / height) * h));
    const red = Number.parseInt(shape.color.slice(0, 2), 16) || 232;
    const green = Number.parseInt(shape.color.slice(2, 4), 16) || 238;
    const blue = Number.parseInt(shape.color.slice(4, 6), 16) || 247;
    const shapePixels = Math.max(0, ex - sx) * Math.max(0, ey - sy);
    if (paintedPixels + shapePixels > MAX_PREVIEW_PAINT_PIXELS) {
      break;
    }
    paintedPixels += shapePixels;
    for (let y = sy; y < ey; y += 1) {
      const row = y * (w * 4 + 1);
      for (let x = sx; x < ex; x += 1) {
        const index = row + 1 + x * 4;
        pixels[index] = red;
        pixels[index + 1] = green;
        pixels[index + 2] = blue;
      }
    }
  }
  const compressed = new Uint8Array(2 + pixels.length + Math.ceil(pixels.length / 65_535) * 5 + 4);
  compressed[0] = 0x78;
  compressed[1] = 0x01;
  let cursor = 2;
  for (let offset = 0; offset < pixels.length;) {
    const length = Math.min(65_535, pixels.length - offset);
    compressed[cursor++] = offset + length === pixels.length ? 1 : 0;
    compressed[cursor++] = length & 255;
    compressed[cursor++] = length >>> 8;
    compressed[cursor++] = ~length & 255;
    compressed[cursor++] = (~length >>> 8) & 255;
    compressed.set(pixels.subarray(offset, offset + length), cursor);
    cursor += length;
    offset += length;
  }
  // `compressed` was sized for exactly these four trailing bytes, so the
  // stream is finished in place rather than copied into a second buffer the
  // same size as the raster.
  new DataView(compressed.buffer).setUint32(cursor, adler32(pixels));
  const output = compressed;
  const signature = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  const header = new Uint8Array(13);
  new DataView(header.buffer).setUint32(0, w);
  new DataView(header.buffer).setUint32(4, h);
  header[8] = 8;
  header[9] = 6;
  const chunks = [
    signature,
    pngChunk("IHDR", header),
    pngChunk("IDAT", output),
    pngChunk("IEND", new Uint8Array()),
  ];
  const png = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
  let position = 0;
  for (const chunk of chunks) {
    png.set(chunk, position);
    position += chunk.length;
  }
  return png;
};

/** The PNG a descriptor describes, built now because something is about to paint it. */
export const rasterizePreview = (descriptor: PreviewDescriptor): Uint8Array =>
  previewPng(
    descriptor.extent.width,
    descriptor.extent.height,
    descriptor.shapes,
    descriptor.pixelWidth,
    descriptor.pixelHeight,
  );
