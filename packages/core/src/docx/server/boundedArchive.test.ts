import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { DocxArchiveError, loadDocxArchive } from "./boundedArchive";

const makeZip = async (entries: Record<string, string | Uint8Array>): Promise<Uint8Array> => {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(entries)) {
    zip.file(path, content);
  }
  return await zip.generateAsync({ type: "uint8array" });
};

/** JSZip reports this capability at runtime; the published typings omit it. */
const NODE_STREAM_SUPPORT = "nodestream";

/**
 * Run `work` with JSZip reporting the capabilities of a browser or web worker,
 * where `readable-stream` is absent and `nodeStream` throws.
 */
const withoutNodeStreamSupport = async <T>(work: () => Promise<T>): Promise<T> => {
  const supported: unknown = Reflect.get(JSZip.support, NODE_STREAM_SUPPORT);
  Reflect.set(JSZip.support, NODE_STREAM_SUPPORT, false);
  try {
    return await work();
  } finally {
    Reflect.set(JSZip.support, NODE_STREAM_SUPPORT, supported);
  }
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
    expect(archive.entryMetadata).toEqual([
      { path: "word/", directory: true, declaredUncompressedBytes: null },
      {
        path: "word/document.xml",
        directory: false,
        declaredUncompressedBytes: 7,
      },
      { path: "media.bin", directory: false, declaredUncompressedBytes: 3 },
    ]);
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

  test("rejects excessive input size before archive parsing", async () => {
    const error = await rejection(loadDocxArchive(new Uint8Array(5), { maxInputBytes: 4 }));

    expect(error).toMatchObject({ reason: "input-too-large" });
  });

  test("rejects invalid archive and per-read limits", async () => {
    const bytes = await makeZip({ entry: "content" });
    const archiveErrors = await Promise.all(
      [Number.NaN, -1, 1.5].map((maxEntryBytes) =>
        rejection(loadDocxArchive(bytes, { maxEntryBytes })),
      ),
    );
    const archive = await loadDocxArchive(bytes);
    const readErrors = await Promise.all(
      [Number.NaN, -1, 1.5].map((maxBytes) =>
        rejection(archive.readEntryUint8("entry", { maxBytes })),
      ),
    );

    expect(archiveErrors).toEqual(
      archiveErrors.map(() => expect.objectContaining({ reason: "invalid-options" })),
    );
    expect(readErrors).toEqual(
      readErrors.map(() => expect.objectContaining({ reason: "invalid-options" })),
    );
  });

  test("rejects declared entry and cumulative sizes", async () => {
    const bytes = await makeZip({ a: "12345", b: "67890" });
    const entryError = await rejection(loadDocxArchive(bytes, { maxEntryBytes: 4 }));
    const totalError = await rejection(loadDocxArchive(bytes, { maxTotalBytes: 8 }));

    expect(entryError).toMatchObject({ reason: "entry-too-large" });
    expect(totalError).toMatchObject({ reason: "total-too-large" });
  });

  test("counts files after directory entries during declared-size preflight", async () => {
    const bytes = await makeZip({ "word/a": "12345", "word/b": "67890" });

    const error = await rejection(loadDocxArchive(bytes, { maxTotalBytes: 8 }));

    expect(error).toMatchObject({ reason: "total-too-large" });
  });

  test("applies a stricter byte limit to an individual read", async () => {
    const archive = await loadDocxArchive(await makeZip({ large: "12345", small: "123" }));

    await expect(archive.readEntryUint8("large", { maxBytes: 4 })).rejects.toMatchObject({
      reason: "entry-too-large",
    });
    await expect(archive.readEntryUint8("small", { maxBytes: 4 })).resolves.toEqual(
      new Uint8Array([49, 50, 51]),
    );
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

  test("keeps a UTF-8 BOM in decoded part text", async () => {
    // Word writes a BOM on OOXML parts. Callers splice part text and write it
    // back, so decoding must not silently drop it.
    const bytes = await makeZip({
      "word/document.xml": new Uint8Array([0xef, 0xbb, 0xbf, 0x3c, 0x61, 0x2f, 0x3e]),
    });
    const archive = await loadDocxArchive(bytes);

    expect(await archive.readEntryString("word/document.xml")).toBe("\uFEFF<a/>");
  });
});

describe("loadDocxArchive on a platform without Node streams", () => {
  test("the simulation removes the Node stream path", async () => {
    const zip = await JSZip.loadAsync(await makeZip({ entry: "content" }));
    const entry = zip.file("entry");
    if (!entry) {
      throw new Error("Expected the fixture entry to exist");
    }

    await withoutNodeStreamSupport(async () => {
      expect(() => entry.nodeStream()).toThrow("nodestream is not supported by this platform");
    });
  });

  test("reads text and bytes and still enforces the entry cap", async () => {
    await withoutNodeStreamSupport(async () => {
      const archive = await loadDocxArchive(await makeZip({ large: "12345", small: "123" }));

      expect(await archive.readEntryString("large")).toBe("12345");
      expect(await archive.readEntryUint8("small")).toEqual(new Uint8Array([49, 50, 51]));
      await expect(archive.readEntryUint8("large", { maxBytes: 4 })).rejects.toMatchObject({
        reason: "entry-too-large",
      });
    });
  });

  test("still enforces the cumulative cap across reads", async () => {
    await withoutNodeStreamSupport(async () => {
      const archive = await loadDocxArchive(await makeZip({ a: "12345", b: "678" }), {
        maxTotalBytes: 8,
      });

      expect(await archive.readEntryString("a")).toBe("12345");
      expect(await archive.readEntryString("b")).toBe("678");
      await expect(archive.readEntryString("a")).rejects.toMatchObject({
        reason: "total-too-large",
      });
    });
  });
});
