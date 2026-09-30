import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { manyEntryArchive, understatedEntryArchive } from "./__tests__/archiveInflationFixtures";
import {
  countCentralDirectoryRecords,
  createInflationBudget,
  DOCX_MAX_COMPRESSION_RATIO,
  inflateEntryWithinLimits,
} from "./archiveInflation";

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
});
