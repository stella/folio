import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { DocxArchiveError, loadDocxArchive } from "./boundedArchive";

const makeZip = async (entries: Record<string, string>): Promise<Uint8Array> => {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(entries)) {
    zip.file(path, content);
  }
  return await zip.generateAsync({ type: "uint8array" });
};

const rejection = async (promise: Promise<unknown>): Promise<unknown> => {
  try {
    await promise;
  } catch (error) {
    return error;
  }
  throw new Error("Expected promise to reject");
};

describe("loadDocxArchive", () => {
  test("reads text and bytes and reports missing entries", async () => {
    const archive = await loadDocxArchive(
      await makeZip({ "word/document.xml": "content", "media.bin": "abc" }),
    );

    expect(await archive.readEntryString("word/document.xml")).toBe("content");
    expect(await archive.readEntryUint8("media.bin")).toEqual(new Uint8Array([97, 98, 99]));
    expect(await archive.readEntryString("missing.xml")).toBeNull();
  });

  test("rejects invalid archives with a tagged error", async () => {
    const error = await rejection(loadDocxArchive(new TextEncoder().encode("not a zip")));

    expect(error).toBeInstanceOf(DocxArchiveError);
    expect(error).toMatchObject({
      _tag: "DocxArchiveError",
      reason: "load-failed",
    });
  });

  test("rejects excessive entry counts before reading", async () => {
    const error = await rejection(
      loadDocxArchive(await makeZip({ a: "1", b: "2" }), {
        maxEntries: 1,
      }),
    );

    expect(error).toMatchObject({ reason: "too-many-entries" });
  });

  test("rejects declared entry and cumulative sizes", async () => {
    const bytes = await makeZip({ a: "12345", b: "67890" });
    const entryError = await rejection(loadDocxArchive(bytes, { maxEntryBytes: 4 }));
    const totalError = await rejection(loadDocxArchive(bytes, { maxTotalBytes: 8 }));

    expect(entryError).toMatchObject({ reason: "entry-too-large" });
    expect(totalError).toMatchObject({ reason: "total-too-large" });
  });

  test("serializes concurrent reads against the cumulative budget", async () => {
    const archive = await loadDocxArchive(await makeZip({ a: "12345", b: "67890" }), {
      maxEntryBytes: 5,
      maxTotalBytes: 10,
    });

    const results = await Promise.allSettled([
      archive.readEntryString("a"),
      archive.readEntryString("b"),
      archive.readEntryString("a"),
    ]);

    expect(results.map(({ status }) => status)).toEqual(["fulfilled", "fulfilled", "rejected"]);
    expect(results.at(2)).toMatchObject({
      reason: { reason: "total-too-large" },
    });
  });
});
