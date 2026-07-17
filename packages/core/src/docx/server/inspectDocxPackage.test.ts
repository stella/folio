import { createHash } from "node:crypto";

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createEmptyDocx } from "../rezip";
import { DocxArchiveError } from "./boundedArchive";
import {
  FOLIO_DOCX_PACKAGE_INSPECTION_VERSION,
  FolioDocxPackageInspectionError,
  inspectDocxPackage,
  readRequestedXmlParts,
  type FolioDocxPackagePart,
} from "./inspectDocxPackage";

type MutatePackage = (zip: JSZip) => void;

const mutateEmptyPackage = async (mutate: MutatePackage): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  mutate(zip);
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

describe("inspectDocxPackage", () => {
  test("lists deterministic part metadata without returning bodies by default", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file("word/media/payload.bin", new Uint8Array(1024));
    });

    const inspection = await inspectDocxPackage(bytes, {
      limits: { maxXmlPartBytes: 4 },
    });

    expect(inspection.version).toBe(FOLIO_DOCX_PACKAGE_INSPECTION_VERSION);
    expect(inspection.xmlParts).toEqual([]);
    expect(inspection.parts.map(({ path }) => path)).toEqual(
      inspection.parts.map(({ path }) => path).toSorted(),
    );
    expect(inspection.parts.find(({ path }) => path === "word/document.xml")).toMatchObject({
      kind: "xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    });
    expect(inspection.parts.find(({ path }) => path === "word/media/payload.bin")).toMatchObject({
      kind: "binary",
      declaredUncompressedBytes: 1024,
    });
  });

  test("returns only explicitly requested XML with an exact byte hash", async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx());
    const documentBytes = await zip.file("word/document.xml")?.async("uint8array");
    if (documentBytes === undefined) {
      throw new Error("Expected document part");
    }

    const inspection = await inspectDocxPackage(await zip.generateAsync({ type: "uint8array" }), {
      xmlParts: ["word/document.xml"],
    });

    expect(inspection.xmlParts).toHaveLength(1);
    expect(inspection.xmlParts.at(0)).toMatchObject({
      path: "word/document.xml",
      byteLength: documentBytes.byteLength,
      sha256: createHash("sha256").update(documentBytes).digest("hex"),
    });
    expect(inspection.xmlParts.at(0)?.text).toContain("<w:document");
  });

  test("decodes UTF-16 XML part bodies", async () => {
    const text = '<?xml version="1.0" encoding="UTF-16"?><root>Žluťoučký</root>';
    const encoded = Buffer.from(text, "utf16le");
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file("custom.xml", Buffer.concat([Buffer.from([0xff, 0xfe]), encoded]));
    });

    const inspection = await inspectDocxPackage(bytes, { xmlParts: ["custom.xml"] });

    expect(inspection.xmlParts.at(0)?.text).toBe(text);
  });

  test("decodes empty and one-byte XML parts", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file("empty.xml", new Uint8Array());
      zip.file("short.xml", "<");
    });

    const inspection = await inspectDocxPackage(bytes, {
      xmlParts: ["empty.xml", "short.xml"],
    });

    expect(inspection.xmlParts.map(({ text }) => text)).toEqual(["", "<"]);
  });

  test("rejects missing, binary, and duplicate requested parts", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file("word/media/payload.bin", "data");
    });

    const missing = await rejection(inspectDocxPackage(bytes, { xmlParts: ["missing.xml"] }));
    const binary = await rejection(
      inspectDocxPackage(bytes, { xmlParts: ["word/media/payload.bin"] }),
    );
    const duplicate = await rejection(
      inspectDocxPackage(bytes, { xmlParts: ["word/document.xml", "word/document.xml"] }),
    );

    expect(missing).toMatchObject({ code: "part-not-found" });
    expect(binary).toMatchObject({ code: "part-not-xml" });
    expect(duplicate).toMatchObject({ code: "duplicate-xml-part" });
    expect(missing).toBeInstanceOf(FolioDocxPackageInspectionError);
  });

  test("enforces requested part count and byte budgets", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file("one.xml", "<one/>");
      zip.file("two.xml", "<two/>");
    });

    const count = await rejection(
      inspectDocxPackage(bytes, {
        xmlParts: ["one.xml", "two.xml"],
        limits: { maxXmlParts: 1 },
      }),
    );
    const partBytes = await rejection(
      inspectDocxPackage(bytes, {
        xmlParts: ["one.xml"],
        limits: { maxXmlPartBytes: 5 },
      }),
    );
    const totalBytes = await rejection(
      inspectDocxPackage(bytes, {
        xmlParts: ["one.xml", "two.xml"],
        limits: { maxXmlTotalBytes: 11 },
      }),
    );

    expect(count).toMatchObject({ code: "too-many-xml-parts" });
    expect(partBytes).toMatchObject({ code: "xml-part-too-large", part: "one.xml" });
    expect(totalBytes).toMatchObject({ code: "xml-total-too-large" });
  });

  test("constrains each XML read to the remaining cumulative budget", async () => {
    const readLimits: number[] = [];
    const part = {
      path: "one.xml",
      kind: "xml",
      contentType: "application/xml",
      declaredUncompressedBytes: null,
    } as const satisfies FolioDocxPackagePart;
    const partsByPath = new Map([
      ["one.xml", part],
      ["two.xml", { ...part, path: "two.xml" }],
    ]);
    const error = await rejection(
      readRequestedXmlParts({
        paths: ["one.xml", "two.xml"],
        partsByPath,
        limits: { maxXmlParts: 2, maxXmlPartBytes: 10, maxXmlTotalBytes: 12 },
        readEntryUint8: (_path, { maxBytes }) => {
          readLimits.push(maxBytes);
          if (readLimits.length === 1) {
            return Promise.resolve(new Uint8Array(6));
          }
          return Promise.reject(
            new DocxArchiveError({
              message: "Read exceeded its limit",
              reason: "entry-too-large",
            }),
          );
        },
      }),
    );

    expect(readLimits).toEqual([10, 6]);
    expect(error).toMatchObject({ code: "xml-total-too-large", part: "two.xml" });
  });

  test("rejects invalid limit values before loading the archive", async () => {
    const error = await rejection(
      inspectDocxPackage(new Uint8Array(), { limits: { maxXmlParts: -1 } }),
    );

    expect(error).toMatchObject({ code: "invalid-limits" });
  });
});
