import { describe, expect, test } from "bun:test";

import { ZipArchive, ZipError, type ZipErrorReason } from "./archive";
import { streamZipCodec, webZipCodec, type ZipCodec } from "./codec";

const MEBIBYTE = 1024 * 1024;

const build = async (
  entries: Record<string, string | Uint8Array>,
  compression: "store" | "deflate" = "deflate",
): Promise<Uint8Array<ArrayBuffer>> => {
  const archive = new ZipArchive();
  for (const [path, data] of Object.entries(entries)) archive.set(path, data);
  return await archive.generate({ compression });
};

const reasonOf = async (
  work: () => Promise<unknown>,
): Promise<ZipErrorReason | "accepted"> => {
  try {
    await work();
    return "accepted";
  } catch (error) {
    if (error instanceof ZipError) return error.reason;
    throw error;
  }
};

const readAll = async (bytes: Uint8Array, codec?: ZipCodec): Promise<void> => {
  const archive = await ZipArchive.read(
    bytes,
    codec === undefined ? {} : { codec },
  );
  for (const entry of archive.entries()) await entry.bytes();
};

type Layout = {
  view: DataView;
  end: number;
  directory: number;
  local: number;
};

/** Offsets of the first entry's records in a single-entry archive. */
const layout = (bytes: Uint8Array): Layout => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const end = bytes.byteLength - 22;
  return { view, end, directory: view.getUint32(end + 16, true), local: 0 };
};

describe("declared sizes and ratios", () => {
  test("an entry larger than the entry limit is refused before inflating", async () => {
    const bytes = await build({
      "word/document.xml": "a".repeat(2 * MEBIBYTE),
    });
    expect(
      await reasonOf(() =>
        ZipArchive.read(bytes, { limits: { maxEntryBytes: MEBIBYTE } }),
      ),
    ).toBe("entry-too-large");
  });

  test("entries that together pass the total limit are refused", async () => {
    const bytes = await build({
      "a.xml": "a".repeat(600_000),
      "b.xml": "b".repeat(600_000),
    });
    expect(
      await reasonOf(() =>
        ZipArchive.read(bytes, { limits: { maxTotalBytes: MEBIBYTE } }),
      ),
    ).toBe("total-too-large");
  });

  test("a highly compressed entry past the grace size fails the ratio cap", async () => {
    const bomb = await build({
      "word/document.xml": new Uint8Array(8 * MEBIBYTE),
    });
    expect(await reasonOf(() => ZipArchive.read(bomb))).toBe(
      "compression-ratio",
    );
    // An exempt entry is still held to the archive-wide ratio.
    expect(
      await reasonOf(() =>
        ZipArchive.read(bomb, { isCompressionRatioExempt: () => true }),
      ),
    ).toBe("compression-ratio");
    // Under the grace size the same ratio is accepted.
    const small = await build({
      "word/document.xml": new Uint8Array(MEBIBYTE),
    });
    expect(await reasonOf(() => readAll(small))).toBe("accepted");
  });

  test("too many records are refused from the declared count", async () => {
    const entries: Record<string, string> = {};
    for (let index = 0; index < 11; index++) entries[`part${index}.xml`] = "x";
    const bytes = await build(entries);
    expect(
      await reasonOf(() =>
        ZipArchive.read(bytes, { limits: { maxEntries: 10 } }),
      ),
    ).toBe("too-many-entries");
  });

  for (const [label, codec] of [
    ["platform", undefined],
    ["web", webZipCodec()],
    ["gzip", streamZipCodec("gzip")],
  ] as const) {
    test(`${label}: data that inflates past its declared size is cut off`, async () => {
      // Declare 100 bytes for an entry that inflates to 4096.
      const bytes = await build({ "a.xml": "a".repeat(4096) });
      const { view, directory } = layout(bytes);
      view.setUint32(directory + 24, 100, true);
      view.setUint32(22, 100, true);
      expect(await reasonOf(() => readAll(bytes, codec))).toBe(
        "declared-size-exceeded",
      );
    });

    test(`${label}: data that inflates short of its declared size is refused`, async () => {
      const bytes = await build({ "a.xml": "a".repeat(4096) });
      const { view, directory } = layout(bytes);
      view.setUint32(directory + 24, 5000, true);
      view.setUint32(22, 5000, true);
      expect(await reasonOf(() => readAll(bytes, codec))).toBe("corrupt-entry");
    });

    test(`${label}: a CRC-32 mismatch is refused`, async () => {
      const bytes = await build({ "a.xml": "abc" }, "store");
      // Flip one stored data byte: sizes still agree, the checksum does not.
      // SAFETY: the stored data begins after the 30-byte header and the name.
      bytes[30 + "a.xml".length] = 0x7a;
      expect(await reasonOf(() => readAll(bytes, codec))).toBe("corrupt-entry");
    });
  }
});

describe("structure", () => {
  test("names that escape the archive root are refused", async () => {
    for (const name of [
      "../evil.xml",
      "/abs.xml",
      "a\\b.xml",
      "a/../../b.xml",
      "C:/x.xml",
    ]) {
      const placeholder = "x".repeat(name.length);
      const bytes = await build({ [placeholder]: "x" });
      const encoded = new TextEncoder().encode(name);
      const { directory } = layout(bytes);
      bytes.set(encoded, 30);
      bytes.set(encoded, directory + 46);
      expect(await reasonOf(() => ZipArchive.read(bytes))).toBe(
        "invalid-entry-name",
      );
    }
  });

  test("a name that is not UTF-8 is refused", async () => {
    const bytes = await build({ ab: "x" });
    const { directory } = layout(bytes);
    bytes.set([0xff, 0xfe], 30);
    bytes.set([0xff, 0xfe], directory + 46);
    expect(await reasonOf(() => ZipArchive.read(bytes))).toBe(
      "invalid-entry-name",
    );
  });

  test("a duplicate name is refused", async () => {
    const bytes = await build({ aa: "x", ab: "y" });
    const view = new DataView(bytes.buffer);
    const directory = view.getUint32(bytes.byteLength - 22 + 16, true);
    const second = directory + 46 + 2;
    const secondLocal = view.getUint32(second + 42, true);
    bytes[second + 46 + 1] = 0x61;
    bytes[secondLocal + 30 + 1] = 0x61;
    expect(await reasonOf(() => ZipArchive.read(bytes))).toBe(
      "duplicate-entry",
    );
  });

  test("a local header that disagrees with the directory is refused", async () => {
    for (const mutate of [
      // name
      (bytes: Uint8Array) => {
        bytes[30] = 0x62;
      },
      // method
      (bytes: Uint8Array) => {
        new DataView(bytes.buffer).setUint16(8, 0, true);
      },
      // crc
      (bytes: Uint8Array) => {
        new DataView(bytes.buffer).setUint32(14, 1, true);
      },
    ]) {
      const bytes = await build({ "a.xml": "a".repeat(100) });
      mutate(bytes);
      expect(await reasonOf(() => ZipArchive.read(bytes))).toBe("malformed");
    }
  });

  test("entries that share bytes are refused", async () => {
    const bytes = await build({ aa: "x".repeat(50), ab: "y".repeat(50) });
    const view = new DataView(bytes.buffer);
    const directory = view.getUint32(bytes.byteLength - 22 + 16, true);
    // Point the second record at the first entry's local header, renamed to match.
    const second = directory + 46 + 2;
    view.setUint32(second + 42, 0, true);
    bytes[second + 46 + 1] = 0x61;
    // Its name now duplicates the first; give the first a different one.
    bytes[directory + 46 + 1] = 0x63;
    bytes[30 + 1] = 0x61;
    expect(["malformed", "duplicate-entry"]).toContain(
      await reasonOf(() => ZipArchive.read(bytes)),
    );
  });

  test("encryption and unknown methods are refused", async () => {
    const encrypted = await build({ "a.xml": "x" });
    const { view, directory } = layout(encrypted);
    view.setUint16(directory + 8, 1, true);
    view.setUint16(6, 1, true);
    expect(await reasonOf(() => ZipArchive.read(encrypted))).toBe(
      "unsupported",
    );

    const bzip2 = await build({ "a.xml": "x" });
    const bz = layout(bzip2);
    bz.view.setUint16(bz.directory + 10, 12, true);
    bz.view.setUint16(8, 12, true);
    expect(await reasonOf(() => ZipArchive.read(bzip2))).toBe("unsupported");
  });

  test("a truncated archive is refused", async () => {
    const bytes = await build({ "a.xml": "a".repeat(100), "b.xml": "b" });
    for (const length of [0, 10, bytes.byteLength - 1, bytes.byteLength - 30]) {
      expect(
        await reasonOf(() => ZipArchive.read(bytes.subarray(0, length))),
      ).toBe("malformed");
    }
  });

  test("an end record that miscounts its records is refused", async () => {
    const bytes = await build({ "a.xml": "x", "b.xml": "y" });
    new DataView(bytes.buffer).setUint16(bytes.byteLength - 22 + 10, 1, true);
    new DataView(bytes.buffer).setUint16(bytes.byteLength - 22 + 8, 1, true);
    expect(await reasonOf(() => ZipArchive.read(bytes))).toBe("malformed");
  });
});
