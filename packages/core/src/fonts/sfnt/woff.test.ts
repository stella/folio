import { describe, expect, test } from "bun:test";
import { deflateSync } from "node:zlib";

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

const WOFF_HEADER_SIZE = 44;
const WOFF_ENTRY_SIZE = 20;

type HandBuiltWoffOptions = {
  readonly payload: Uint8Array;
  readonly numTables: number;
  /** Decompressed size each entry claims, whatever the payload really holds. */
  readonly origLength: number;
};

/**
 * A WOFF container whose every directory entry names the same stored span:
 * the shape a per-table ceiling does not bound, because one payload can be
 * charged to the decoder as many times as the directory has room for.
 */
const woffSharingOnePayload = ({
  payload,
  numTables,
  origLength,
}: HandBuiltWoffOptions): Uint8Array => {
  const dataOffset = WOFF_HEADER_SIZE + numTables * WOFF_ENTRY_SIZE;
  const bytes = new Uint8Array(dataOffset + payload.byteLength);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, SFNT_VERSION.woff);
  view.setUint32(4, SFNT_VERSION.trueType);
  view.setUint32(8, bytes.byteLength);
  view.setUint16(12, numTables);
  view.setUint32(WOFF_TOTAL_SFNT_SIZE_OFFSET, numTables * origLength);

  for (let index = 0; index < numTables; index++) {
    const entry = WOFF_HEADER_SIZE + index * WOFF_ENTRY_SIZE;
    const tag = `t${String(index).padStart(3, "0")}`;
    for (let character = 0; character < tag.length; character++) {
      view.setUint8(entry + character, tag.charCodeAt(character));
    }
    view.setUint32(entry + 4, dataOffset);
    view.setUint32(entry + 8, payload.byteLength);
    view.setUint32(entry + 12, origLength);
    view.setUint32(entry + 16, 0);
  }
  bytes.set(payload, dataOffset);
  return bytes;
};

describe("decodeWoff resource bounds", () => {
  /**
   * A size each entry may legally claim on its own, so only the container-wide
   * total can refuse the file. Should the per-table ceiling ever drop below
   * this, the assertion on the message below is what says so.
   */
  const PER_TABLE_LEGAL_LENGTH = 1 << 27;

  test("refuses a directory whose entries share one span to claim the whole heap", () => {
    const table = new Uint8Array(64).fill(0x2a);
    const deflated = deflateSync(table);
    const woff = woffSharingOnePayload({
      payload: new Uint8Array(deflated.buffer, deflated.byteOffset, deflated.byteLength),
      numTables: 8,
      origLength: PER_TABLE_LEGAL_LENGTH,
    });

    const decoded = decodeWoff(woff);
    expect(decoded.isErr()).toBe(true);
    if (decoded.isErr()) {
      // Rejected on the declared total, before any entry is inflated: an
      // error about one table's inflated size means the container-wide
      // ceiling never ran.
      expect(decoded.error.message).toContain("in total");
    }
  });

  test("still decodes a container whose tables add up to an ordinary font", () => {
    const table = new Uint8Array(256).fill(0x2a);
    const deflated = deflateSync(table);
    const woff = woffSharingOnePayload({
      payload: new Uint8Array(deflated.buffer, deflated.byteOffset, deflated.byteLength),
      numTables: 2,
      origLength: table.byteLength,
    });

    const decoded = decodeWoff(woff);
    expect(decoded.isErr()).toBe(false);
    if (decoded.isErr()) {
      return;
    }
    const directory = readSfntDirectory(decoded.value);
    expect(directory.isErr()).toBe(false);
    if (directory.isErr()) {
      return;
    }
    expect([...directory.value.tables.keys()]).toEqual(["t000", "t001"]);
    for (const record of directory.value.tables.values()) {
      expect(record.length).toBe(table.byteLength);
      expect([...decoded.value.subarray(record.offset, record.offset + record.length)]).toEqual([
        ...table,
      ]);
    }
  });
});
