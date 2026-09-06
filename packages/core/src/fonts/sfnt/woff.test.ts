import { describe, expect, test } from "bun:test";

import { parseSfnt } from "./parse";
import { readSfntDirectory, SFNT_VERSION } from "./tables";
import {
  TEST_FONT_FAMILIES,
  TEST_FONTS_INSTALLED,
  TEST_FONTS_SKIP_REASON,
  readTestFont,
} from "./__tests__/testFonts";
import { decodeWoff, toSfntBytes } from "./woff";

/** Offset of `totalSfntSize` in the WOFF header: the size we must reproduce. */
const WOFF_TOTAL_SFNT_SIZE_OFFSET = 16;

const magic = (tag: string): Uint8Array => {
  const bytes = new Uint8Array(64);
  for (let index = 0; index < tag.length; index++) {
    bytes[index] = tag.charCodeAt(index);
  }
  return bytes;
};

describe(`decodeWoff (${TEST_FONTS_SKIP_REASON})`, () => {
  for (const family of TEST_FONT_FAMILIES) {
    test.skipIf(!TEST_FONTS_INSTALLED)(`${family} decodes into a parseable sfnt`, async () => {
      const woff = await readTestFont({ family });
      const decoded = decodeWoff(woff);

      expect(decoded.isErr()).toBe(false);
      if (decoded.isErr()) {
        return;
      }

      const sfnt = decoded.value;
      const view = new DataView(sfnt.buffer, sfnt.byteOffset, sfnt.byteLength);
      expect(view.getUint32(0)).toBe(SFNT_VERSION.trueType);

      // The container records the exact size of the font it was built from;
      // rebuilding it any other size means the padding or directory is wrong.
      const woffView = new DataView(woff.buffer, woff.byteOffset, woff.byteLength);
      expect(sfnt.byteLength).toBe(woffView.getUint32(WOFF_TOTAL_SFNT_SIZE_OFFSET));

      const directory = readSfntDirectory(sfnt);
      expect(directory.isErr()).toBe(false);
      if (directory.isErr()) {
        return;
      }
      expect(directory.value.tables.size).toBeGreaterThan(4);
      const tags = [...directory.value.tables.keys()];
      expect(tags).toContain("head");
      expect(tags).toContain("glyf");
      // Every table starts on a four-byte boundary.
      for (const table of directory.value.tables.values()) {
        expect(table.offset % 4).toBe(0);
      }

      const font = parseSfnt(sfnt);
      expect(font.isErr()).toBe(false);
    });
  }

  test.skipIf(!TEST_FONTS_INSTALLED)("is deterministic across runs", async () => {
    const woff = await readTestFont({ family: "arimo" });
    const first = decodeWoff(woff);
    const second = decodeWoff(woff);
    expect(first.isErr()).toBe(false);
    expect(second.isErr()).toBe(false);
    if (first.isErr() || second.isErr()) {
      return;
    }
    expect(Buffer.from(first.value).equals(Buffer.from(second.value))).toBe(true);
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("rejects a truncated container", async () => {
    const woff = await readTestFont({ family: "arimo" });
    const decoded = decodeWoff(woff.subarray(0, 200));
    expect(decoded.isErr()).toBe(true);
  });

  test("rejects a buffer that is not WOFF", () => {
    const decoded = decodeWoff(magic("OTTO"));
    expect(decoded.isErr()).toBe(true);
    if (decoded.isErr()) {
      expect(decoded.error.message).toContain("not a WOFF container");
    }
  });

  test("rejects a header shorter than 44 bytes", () => {
    const decoded = decodeWoff(new Uint8Array(12));
    expect(decoded.isErr()).toBe(true);
  });
});

describe("toSfntBytes", () => {
  test.skipIf(!TEST_FONTS_INSTALLED)("passes plain sfnt bytes through", async () => {
    const woff = await readTestFont({ family: "arimo" });
    const decoded = decodeWoff(woff);
    expect(decoded.isErr()).toBe(false);
    if (decoded.isErr()) {
      return;
    }

    const passed = toSfntBytes(decoded.value);
    expect(passed.isErr()).toBe(false);
    if (passed.isErr()) {
      return;
    }
    expect(Buffer.from(passed.value).equals(Buffer.from(decoded.value))).toBe(true);
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("decodes a WOFF container", async () => {
    const woff = await readTestFont({ family: "tinos" });
    const converted = toSfntBytes(woff);
    expect(converted.isErr()).toBe(false);
    if (converted.isErr()) {
      return;
    }
    expect(parseSfnt(converted.value).isErr()).toBe(false);
  });

  test("rejects WOFF2 by name", () => {
    const converted = toSfntBytes(magic("wOF2"));
    expect(converted.isErr()).toBe(true);
    if (converted.isErr()) {
      expect(converted.error.message).toContain("WOFF2");
    }
  });

  test("rejects an unknown container", () => {
    const converted = toSfntBytes(magic("RIFF"));
    expect(converted.isErr()).toBe(true);
  });

  test("rejects a buffer too short to identify", () => {
    expect(toSfntBytes(new Uint8Array(2)).isErr()).toBe(true);
  });
});
