import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import {
  archiveWithRawDeflate,
  manyEntryArchive,
  rawDeflateOfZeros,
  understatedEntryArchive,
} from "./__tests__/archiveInflationFixtures";
import {
  compressionRatioLimitFor,
  countCentralDirectoryRecords,
  createInflationBudget,
  DOCX_MAX_COMPRESSION_RATIO,
  inflateEntryWithinLimits,
  isRatioBoundedPart,
  isStoredZipEntry,
} from "./archiveInflation";

const MEBIBYTE = 1024 * 1024;

/** Count every byte the entry's stream produces, whether or not it is kept. */
const observeInflatedBytes = (entry: JSZip.JSZipObject): (() => number) => {
  let seen = 0;
  const openStream = entry.internalStream.bind(entry);
  entry.internalStream = (type) =>
    openStream(type).on("data", (chunk) => {
      seen += chunk.length;
    });
  return () => seen;
};

const loadEntry = async (bytes: Uint8Array, path: string): Promise<JSZip.JSZipObject> => {
  const entry = (await JSZip.loadAsync(bytes)).file(path);
  if (!entry) {
    throw new Error(`Expected the fixture entry ${path}`);
  }
  return entry;
};

const storedArchive = async (entries: Record<string, string>): Promise<Uint8Array> => {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(entries)) {
    zip.file(path, content, { createFolders: false });
  }
  return await zip.generateAsync({ type: "uint8array" });
};

describe("countCentralDirectoryRecords", () => {
  test("counts every record and stops once past the bound", async () => {
    const bytes = await manyEntryArchive(20);

    expect(countCentralDirectoryRecords(bytes, 1000)).toBe(23);
    expect(countCentralDirectoryRecords(bytes, 5)).toBe(6);
  });
});

describe("inflateEntryWithinLimits", () => {
  test("stops an entry at its declared size and refunds the budget", async () => {
    const entry = await loadEntry(
      await understatedEntryArchive("word/header1.xml"),
      "word/header1.xml",
    );
    const budget = createInflationBudget(Number.MAX_SAFE_INTEGER);

    const result = await inflateEntryWithinLimits({
      entry,
      maxEntryBytes: Number.MAX_SAFE_INTEGER,
      maxCompressionRatio: DOCX_MAX_COMPRESSION_RATIO,
      budget,
    });

    expect(result).toEqual({ ok: false, limit: "declared-size" });
    expect(budget.inflatedBytes).toBe(0);
  });

  test("charges one budget across entries and stops the entry that passes it", async () => {
    const bytes = await storedArchive({ a: "12345", b: "67890" });
    const budget = createInflationBudget(8);
    const inflate = async (path: string) =>
      await inflateEntryWithinLimits({
        entry: await loadEntry(bytes, path),
        maxEntryBytes: 5,
        maxCompressionRatio: DOCX_MAX_COMPRESSION_RATIO,
        budget,
      });

    expect(await inflate("a")).toEqual({ ok: true, bytes: new TextEncoder().encode("12345") });
    expect(await inflate("b")).toEqual({ ok: false, limit: "total" });
    expect(budget.inflatedBytes).toBe(5);
  });

  test("applies the caller's per-entry cap", async () => {
    const bytes = await storedArchive({ a: "12345" });

    const result = await inflateEntryWithinLimits({
      entry: await loadEntry(bytes, "a"),
      maxEntryBytes: 4,
      maxCompressionRatio: DOCX_MAX_COMPRESSION_RATIO,
      budget: createInflationBudget(100),
    });

    expect(result).toEqual({ ok: false, limit: "entry" });
  });

  test("verifies without keeping the bytes", async () => {
    const bytes = await storedArchive({ a: "12345" });
    const budget = createInflationBudget(100);

    const result = await inflateEntryWithinLimits({
      entry: await loadEntry(bytes, "a"),
      maxEntryBytes: 100,
      maxCompressionRatio: DOCX_MAX_COMPRESSION_RATIO,
      budget,
      retain: false,
    });

    expect(result).toEqual({ ok: true, bytes: new Uint8Array(0) });
    expect(budget.inflatedBytes).toBe(5);
  });

  test("stops an entry declared at 64 bytes that would inflate to a gibibyte", async () => {
    // About 1 MiB of DEFLATE data that inflates to 1 GiB of zeros.
    const deflated = rawDeflateOfZeros(1024);
    expect(deflated.byteLength).toBeLessThan(2 * MEBIBYTE);
    const entry = await loadEntry(
      await archiveWithRawDeflate({ entryPath: "zeros.bin", deflated, declaredBytes: 64 }),
      "zeros.bin",
    );
    const inflatedBytes = observeInflatedBytes(entry);
    const budget = createInflationBudget(Number.MAX_SAFE_INTEGER);

    const result = await inflateEntryWithinLimits({
      entry,
      maxEntryBytes: Number.MAX_SAFE_INTEGER,
      maxCompressionRatio: Number.POSITIVE_INFINITY,
      budget,
    });

    expect(result).toEqual({ ok: false, limit: "declared-size" });
    // One compressed chunk can still inflate after the stream is paused; the
    // point is that the rest of the entry never is.
    expect(inflatedBytes()).toBeLessThan(32 * MEBIBYTE);
    expect(budget.inflatedBytes).toBe(0);
  });

  test("refunds only the entry that stops when inflations share a budget", async () => {
    const understated = await understatedEntryArchive("word/header1.xml");
    const budget = createInflationBudget(Number.MAX_SAFE_INTEGER);
    const inflate = async (path: string) =>
      await inflateEntryWithinLimits({
        entry: await loadEntry(understated, path),
        maxEntryBytes: Number.MAX_SAFE_INTEGER,
        maxCompressionRatio: DOCX_MAX_COMPRESSION_RATIO,
        budget,
      });

    const [kept, stopped] = await Promise.all([
      inflate("word/document.xml"),
      inflate("word/header1.xml"),
    ]);

    expect(kept).toEqual({ ok: true, bytes: new TextEncoder().encode("<w:document/>") });
    expect(stopped).toEqual({ ok: false, limit: "declared-size" });
    expect(budget.inflatedBytes).toBe("<w:document/>".length);
  });

  test("stops at the next chunk once the shared budget is aborted", async () => {
    const bytes = await storedArchive({ a: "12345" });
    const budget = createInflationBudget(100);
    budget.aborted = true;

    const result = await inflateEntryWithinLimits({
      entry: await loadEntry(bytes, "a"),
      maxEntryBytes: 100,
      maxCompressionRatio: DOCX_MAX_COMPRESSION_RATIO,
      budget,
    });

    expect(result).toEqual({ ok: false, limit: "aborted" });
    expect(budget.inflatedBytes).toBe(0);
  });
});

describe("entry classification", () => {
  test("tells stored entries from compressed ones", async () => {
    const zip = new JSZip();
    zip.file("stored", "abc", { compression: "STORE" });
    zip.file("compressed", "abc", { compression: "DEFLATE" });
    const bytes = await zip.generateAsync({ type: "uint8array" });

    expect(isStoredZipEntry(await loadEntry(bytes, "stored"))).toBe(true);
    expect(isStoredZipEntry(await loadEntry(bytes, "compressed"))).toBe(false);
  });

  test("applies the per-entry ratio to markup and text parts only", () => {
    expect(isRatioBoundedPart("word/document.xml")).toBe(true);
    expect(isRatioBoundedPart("word/_rels/document.xml.rels")).toBe(true);
    expect(isRatioBoundedPart("word/media/image1.BMP")).toBe(false);
    expect(compressionRatioLimitFor("word/media/image1.emf", 200)).toBe(Number.POSITIVE_INFINITY);
    expect(compressionRatioLimitFor("word/header1.xml", 200)).toBe(200);
  });
});
