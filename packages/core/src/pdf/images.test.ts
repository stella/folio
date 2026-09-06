import { describe, expect, test } from "bun:test";
import { deflateSync } from "node:zlib";
import type { DisplayImageSource } from "../display-list/types";
import { decodeImage } from "./images";

const CRC_POLYNOMIAL = 0xedb88320;

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < table.length; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) === 1 ? CRC_POLYNOMIAL ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value;
  }
  return table;
})();

const crc32 = (bytes: Uint8Array): number => {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc = (crcTable[(crc ^ byte) & 0xff] ?? 0) ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
};

const concat = (parts: readonly Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let cursor = 0;
  for (const part of parts) {
    out.set(part, cursor);
    cursor += part.length;
  }
  return out;
};

const chunk = (type: string, data: Uint8Array): Uint8Array => {
  const header = new Uint8Array(8);
  new DataView(header.buffer).setUint32(0, data.length);
  for (let index = 0; index < 4; index += 1) {
    header[4 + index] = type.charCodeAt(index);
  }
  const crcInput = concat([header.subarray(4), data]);
  const trailer = new Uint8Array(4);
  new DataView(trailer.buffer).setUint32(0, crc32(crcInput));
  return concat([header, data, trailer]);
};

const paeth = (left: number, above: number, upperLeft: number): number => {
  const estimate = left + above - upperLeft;
  const dl = Math.abs(estimate - left);
  const da = Math.abs(estimate - above);
  const du = Math.abs(estimate - upperLeft);
  if (dl <= da && dl <= du) {
    return left;
  }
  return da <= du ? above : upperLeft;
};

/** The forward direction of the five PNG filters, so decoding round-trips. */
const applyFilter = (
  filterType: number,
  row: Uint8Array,
  previous: Uint8Array | null,
  bytesPerPixel: number,
): Uint8Array => {
  const out = new Uint8Array(row.length);
  for (let index = 0; index < row.length; index += 1) {
    const raw = row[index] ?? 0;
    const left = index >= bytesPerPixel ? (row[index - bytesPerPixel] ?? 0) : 0;
    const above = previous?.[index] ?? 0;
    const upperLeft = index >= bytesPerPixel ? (previous?.[index - bytesPerPixel] ?? 0) : 0;
    switch (filterType) {
      case 1:
        out[index] = (raw - left) & 0xff;
        break;
      case 2:
        out[index] = (raw - above) & 0xff;
        break;
      case 3:
        out[index] = (raw - ((left + above) >> 1)) & 0xff;
        break;
      case 4:
        out[index] = (raw - paeth(left, above, upperLeft)) & 0xff;
        break;
      default:
        out[index] = raw;
    }
  }
  return out;
};

type BuildPngOptions = {
  readonly width: number;
  readonly height: number;
  readonly bitDepth: number;
  readonly colorType: number;
  readonly rows: readonly Uint8Array[];
  readonly filters: readonly number[];
  readonly bytesPerPixel: number;
  readonly palette?: Uint8Array;
  readonly transparency?: Uint8Array;
  readonly interlace?: number;
};

const buildPng = ({
  width,
  height,
  bitDepth,
  colorType,
  rows,
  filters,
  bytesPerPixel,
  palette,
  transparency,
  interlace = 0,
}: BuildPngOptions): Uint8Array => {
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  header[8] = bitDepth;
  header[9] = colorType;
  header[12] = interlace;

  const scanlines: Uint8Array[] = [];
  for (const [index, row] of rows.entries()) {
    const filterType = filters[index] ?? 0;
    scanlines.push(new Uint8Array([filterType]));
    scanlines.push(applyFilter(filterType, row, rows[index - 1] ?? null, bytesPerPixel));
  }
  const deflated = deflateSync(concat(scanlines));
  return concat([
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", header),
    ...(palette === undefined ? [] : [chunk("PLTE", palette)]),
    ...(transparency === undefined ? [] : [chunk("tRNS", transparency)]),
    chunk("IDAT", new Uint8Array(deflated.buffer, deflated.byteOffset, deflated.byteLength)),
    chunk("IEND", new Uint8Array(0)),
  ]);
};

const asSource = (bytes: Uint8Array, width: number, height: number): DisplayImageSource => ({
  format: "png",
  bytes,
  pixelWidth: width,
  pixelHeight: height,
});

describe("PNG filter reversal", () => {
  const width = 4;
  const height = 5;
  const bytesPerRow = width * 3;
  const rows = Array.from({ length: height }, (_, row) =>
    Uint8Array.from({ length: bytesPerRow }, (__, index) => (row * 37 + index * 11 + 3) & 0xff),
  );

  test("reverses all five filter types", () => {
    const png = buildPng({
      width,
      height,
      bitDepth: 8,
      colorType: 2,
      rows,
      // One row per filter type, so a defect in any one of them fails here.
      filters: [0, 1, 2, 3, 4],
      bytesPerPixel: 3,
    });
    const decoded = decodeImage(asSource(png, width, height));
    expect(decoded.isErr()).toBe(false);
    if (decoded.isErr()) {
      return;
    }
    expect(decoded.value.widthPx).toBe(width);
    expect(decoded.value.heightPx).toBe(height);
    expect(decoded.value.bitsPerComponent).toBe(8);
    expect([...decoded.value.data]).toEqual([...concat(rows)]);
  });

  test("splits an alpha channel into a soft mask", () => {
    const rgbaRows = Array.from({ length: 2 }, (_, row) =>
      Uint8Array.from({ length: 2 * 4 }, (__, index) => (row * 5 + index * 9 + 1) & 0xff),
    );
    const png = buildPng({
      width: 2,
      height: 2,
      bitDepth: 8,
      colorType: 6,
      rows: rgbaRows,
      filters: [4, 4],
      bytesPerPixel: 4,
    });
    const decoded = decodeImage(asSource(png, 2, 2));
    if (decoded.isErr()) {
      throw decoded.error;
    }
    expect(decoded.value.data.length).toBe(2 * 2 * 3);
    expect(decoded.value.alpha?.data.length).toBe(2 * 2);
    const flat = [...rgbaRows[0]!, ...rgbaRows[1]!];
    expect([...(decoded.value.alpha?.data ?? [])]).toEqual([
      flat[3]!,
      flat[7]!,
      flat[11]!,
      flat[15]!,
    ]);
  });

  test("keeps a sub-byte palette packed and turns tRNS into a soft mask", () => {
    // Four pixels per row at 4 bits each: two bytes a row, indices 0..3.
    const rows4 = [new Uint8Array([0x01, 0x23]), new Uint8Array([0x32, 0x10])];
    const png = buildPng({
      width: 4,
      height: 2,
      bitDepth: 4,
      colorType: 3,
      rows: rows4,
      filters: [0, 1],
      bytesPerPixel: 1,
      palette: new Uint8Array([0, 0, 0, 255, 0, 0, 0, 255, 0, 0, 0, 255]),
      transparency: new Uint8Array([0, 128, 255, 255]),
    });
    const decoded = decodeImage(asSource(png, 4, 2));
    if (decoded.isErr()) {
      throw decoded.error;
    }
    expect(decoded.value.bitsPerComponent).toBe(4);
    expect([...decoded.value.data]).toEqual([0x01, 0x23, 0x32, 0x10]);
    expect([...(decoded.value.alpha?.data ?? [])]).toEqual([0, 128, 255, 255, 255, 255, 128, 0]);
  });
});

describe("PNG forms the writer refuses", () => {
  test("rejects 16-bit samples", () => {
    const png = buildPng({
      width: 1,
      height: 1,
      bitDepth: 16,
      colorType: 2,
      rows: [new Uint8Array(6)],
      filters: [0],
      bytesPerPixel: 6,
    });
    const decoded = decodeImage(asSource(png, 1, 1));
    expect(decoded.isErr()).toBe(true);
  });

  test("rejects an interlaced image", () => {
    const png = buildPng({
      width: 1,
      height: 1,
      bitDepth: 8,
      colorType: 2,
      rows: [new Uint8Array(3)],
      filters: [0],
      bytesPerPixel: 3,
      interlace: 1,
    });
    const decoded = decodeImage(asSource(png, 1, 1));
    expect(decoded.isErr()).toBe(true);
    if (decoded.isErr()) {
      expect(decoded.error.message).toContain("interlaced");
    }
  });
});

describe("JPEG", () => {
  test("passes the bytes through and reads the frame header", () => {
    const jpeg = concat([
      new Uint8Array([0xff, 0xd8]),
      new Uint8Array([0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x40, 0x00, 0x60, 0x03]),
      new Uint8Array(9),
      new Uint8Array([0xff, 0xda, 0x00, 0x02]),
    ]);
    const decoded = decodeImage({ format: "jpeg", bytes: jpeg, pixelWidth: 96, pixelHeight: 64 });
    if (decoded.isErr()) {
      throw decoded.error;
    }
    expect(decoded.value.encoding).toBe("jpeg");
    expect(decoded.value.widthPx).toBe(96);
    expect(decoded.value.heightPx).toBe(64);
    expect(decoded.value.data).toBe(jpeg);
  });

  test("reports a JPEG with no frame header", () => {
    const decoded = decodeImage({
      format: "jpeg",
      bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xda, 0x00, 0x02]),
      pixelWidth: 1,
      pixelHeight: 1,
    });
    expect(decoded.isErr()).toBe(true);
  });
});

describe("PNG streams that lie about their size", () => {
  /** A well-formed 1x1 RGB PNG with its IDAT payload replaced. */
  const withIdat = (payload: Uint8Array): Uint8Array => {
    const png = buildPng({
      width: 1,
      height: 1,
      bitDepth: 8,
      colorType: 2,
      rows: [new Uint8Array([1, 2, 3])],
      filters: [0],
      bytesPerPixel: 3,
    });
    const SIGNATURE_AND_IHDR_BYTES = 8 + 25;
    return concat([
      png.subarray(0, SIGNATURE_AND_IHDR_BYTES),
      chunk("IDAT", payload),
      chunk("IEND", new Uint8Array(0)),
    ]);
  };

  test("reports a corrupt IDAT stream instead of throwing", () => {
    const decoded = decodeImage(asSource(withIdat(new Uint8Array([0x78, 0x9c, 0xff, 0xff])), 1, 1));
    expect(decoded.isErr()).toBe(true);
    if (decoded.isErr()) {
      expect(decoded.error.message).toContain("IDAT");
    }
  });

  test("refuses an IDAT that inflates past the declared geometry", () => {
    // 8 MB of zeroes behind a 1x1 IHDR: unbounded, the whole thing is
    // allocated before any check looks at how big the image claims to be.
    const bomb = deflateSync(new Uint8Array(8 << 20));
    const decoded = decodeImage(
      asSource(withIdat(new Uint8Array(bomb.buffer, bomb.byteOffset, bomb.byteLength)), 1, 1),
    );
    expect(decoded.isErr()).toBe(true);
    if (decoded.isErr()) {
      expect(decoded.error.message).toContain("IDAT");
    }
  });
});

describe("JPEG frame modes /DCTDecode cannot represent", () => {
  /** A minimal JPEG whose only frame header carries `marker`. */
  const decodeFrame = (marker: number) =>
    decodeImage({
      format: "jpeg",
      bytes: concat([
        new Uint8Array([0xff, 0xd8]),
        new Uint8Array([0xff, marker, 0x00, 0x11, 0x08, 0x00, 0x40, 0x00, 0x60, 0x03]),
        new Uint8Array(9),
        new Uint8Array([0xff, 0xda, 0x00, 0x02]),
      ]),
      pixelWidth: 96,
      pixelHeight: 64,
    });

  test("keeps progressive alongside the sequential Huffman modes", () => {
    for (const marker of [0xc0, 0xc1, 0xc2]) {
      const decoded = decodeFrame(marker);
      if (decoded.isErr()) {
        throw decoded.error;
      }
      expect(decoded.value.encoding).toBe("jpeg");
      expect(decoded.value.widthPx).toBe(96);
      expect(decoded.value.heightPx).toBe(64);
    }
  });

  test("names every lossless, arithmetic and hierarchical frame it refuses", () => {
    for (const marker of [0xc3, 0xc5, 0xc6, 0xc7, 0xc8, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf]) {
      const decoded = decodeFrame(marker);
      expect(decoded.isErr()).toBe(true);
      if (decoded.isErr()) {
        expect(decoded.error.message).toContain(`0xff${marker.toString(16)}`);
        expect(decoded.error.message).toContain("/DCTDecode");
      }
    }
  });
});
