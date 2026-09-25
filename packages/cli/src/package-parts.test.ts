import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { buildDocx, CONTRACT_PARAGRAPHS } from "./__tests__/fixtures";
import { checkPackageIntegrity, diffPackages, nextRevisionIdSeed } from "./package-parts";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const rewrite = async (
  bytes: Uint8Array,
  edit: (zip: JSZip) => void,
): Promise<Uint8Array<ArrayBuffer>> => {
  const zip = await JSZip.loadAsync(bytes);
  edit(zip);
  return await zip.generateAsync({ type: "uint8array" });
};

const documentXml = (body: string, prefix = "w"): string =>
  `<?xml version="1.0" encoding="UTF-8"?><${prefix}:document xmlns:${prefix}="${W_NS}" ` +
  `xmlns:rel="${R_NS}"><${prefix}:body>${body}</${prefix}:body></${prefix}:document>`;

describe("diffPackages", () => {
  test("reports added, modified, and removed parts by content", async () => {
    const before = await buildDocx(CONTRACT_PARAGRAPHS);
    const after = await rewrite(before, (zip) => {
      zip.file("word/document.xml", documentXml("<w:p/>"));
      zip.file("word/comments.xml", "<comments/>");
      zip.remove("word/styles.xml");
    });
    const recompressed = await rewrite(before, () => undefined);

    expect((await diffPackages(before, after)).unwrap()).toEqual([
      { part: "word/comments.xml", change: "added" },
      { part: "word/document.xml", change: "modified" },
      { part: "word/styles.xml", change: "removed" },
    ]);
    expect((await diffPackages(before, recompressed)).unwrap()).toEqual([]);
  });
});

describe("nextRevisionIdSeed", () => {
  test("sits above every WordprocessingML id, whatever the prefix", async () => {
    const bytes = await rewrite(await buildDocx([]), (zip) => {
      zip.file(
        "word/document.xml",
        documentXml('<x:p><x:ins x:id="41"/><x:bookmarkStart x:id="7"/></x:p>', "x"),
      );
      zip.file(
        "word/other.xml",
        '<root xmlns:w="urn:not-wordprocessingml"><w:ins w:id="900"/></root>',
      );
    });

    expect((await nextRevisionIdSeed(bytes)).unwrap()).toBe(42);
  });
});

describe("checkPackageIntegrity", () => {
  const changed = [{ part: "word/document.xml", change: "modified" } as const];

  test("accepts a well-formed package", async () => {
    const bytes = await buildDocx(CONTRACT_PARAGRAPHS);

    expect((await checkPackageIntegrity(bytes, changed)).isOk()).toBe(true);
  });

  test("refuses malformed XML, dangling references, and missing targets", async () => {
    const base = await buildDocx(CONTRACT_PARAGRAPHS);
    const malformed = await rewrite(base, (zip) => {
      zip.file("word/document.xml", "<w:document><w:body>");
    });
    const dangling = await rewrite(base, (zip) => {
      zip.file("word/document.xml", documentXml('<w:p><w:hyperlink rel:id="rId99"/></w:p>'));
    });
    const missingTarget = await rewrite(base, (zip) => {
      zip.remove("word/styles.xml");
    });

    for (const [bytes, parts] of [
      [malformed, changed],
      [dangling, changed],
      [missingTarget, [{ part: "word/_rels/document.xml.rels", change: "modified" } as const]],
    ] as const) {
      const result = await checkPackageIntegrity(bytes, parts);
      expect(result.isErr() && result.error.code).toBe("integrity_failed");
    }
  });
});
