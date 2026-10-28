import { describe, expect, test } from "bun:test";

import { ZipArchive, ZipError } from "@stll/docx-core/zip";

import {
  archiveWithRawDeflate,
  manyEntryArchive,
  rawDeflateOfZeros,
  understatedEntryArchive,
} from "./__tests__/archiveInflationFixtures";
import { isRatioBoundedPart } from "./archiveInflation";

const MEBIBYTE = 1024 * 1024;

const reasonOf = async (read: () => Promise<unknown>): Promise<string> => {
  try {
    await read();
    return "accepted";
  } catch (error) {
    if (error instanceof ZipError) return error.reason;
    throw error;
  }
};

describe("archive reads hold entries to their declared sizes", () => {
  test("an entry declared smaller than its data stops at the declared size", async () => {
    const zip = await ZipArchive.read(
      await understatedEntryArchive("word/header1.xml"),
    );
    expect(await zip.entry("word/document.xml")?.text()).toBe("<w:document/>");
    expect(
      await reasonOf(async () => await zip.entry("word/header1.xml")?.bytes()),
    ).toBe("declared-size-exceeded");
  });

  test("an entry declared at 64 bytes that would inflate to a gibibyte stops at 64", async () => {
    // About 1 MiB of DEFLATE data that inflates to 1 GiB of zeros.
    const deflated = rawDeflateOfZeros(1024);
    expect(deflated.byteLength).toBeLessThan(2 * MEBIBYTE);
    const zip = await ZipArchive.read(
      await archiveWithRawDeflate({
        entryPath: "zeros.bin",
        deflated,
        declaredBytes: 64,
      }),
    );
    const started = performance.now();
    expect(
      await reasonOf(async () => await zip.entry("zeros.bin")?.bytes()),
    ).toBe("declared-size-exceeded");
    // Inflating the whole gibibyte takes seconds; stopping at 64 bytes does not.
    expect(performance.now() - started).toBeLessThan(1000);
  });

  test("the record count is checked against the declared total first", async () => {
    expect(
      await reasonOf(
        async () =>
          await ZipArchive.read(await manyEntryArchive(20), {
            limits: { maxEntries: 5 },
          }),
      ),
    ).toBe("too-many-entries");
  });
});

describe("entry classification", () => {
  test("applies the per-entry ratio to markup and text parts only", () => {
    expect(isRatioBoundedPart("word/document.xml")).toBe(true);
    expect(isRatioBoundedPart("word/_rels/document.xml.rels")).toBe(true);
    expect(isRatioBoundedPart("word/media/image1.BMP")).toBe(false);
    expect(isRatioBoundedPart("word/media/image1.emf")).toBe(false);
  });
});
