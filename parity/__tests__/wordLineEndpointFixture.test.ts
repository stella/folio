import { describe, expect, test } from "bun:test";
import path from "node:path";

import { readLineEndpointManifest } from "../lineEndpoints";
import { sha256OfFile } from "../pdfReference";

const FIXTURE_PATH = path.join(import.meta.dir, "..", "fixtures", "word-hyphenation-hanging.docx");
const MANIFEST_PATH = path.join(
  import.meta.dir,
  "..",
  "fixtures",
  "word-hyphenation-hanging.word-lines.json",
);

describe("Word hyphenation and hanging-punctuation baseline", () => {
  test("binds the reviewed Word endpoints to the exact synthetic DOCX", async () => {
    const [manifest, sourceSha256] = await Promise.all([
      readLineEndpointManifest(MANIFEST_PATH),
      sha256OfFile(FIXTURE_PATH),
    ]);

    expect(manifest.source).toEqual({
      fileName: "word-hyphenation-hanging.docx",
      sha256: sourceSha256,
    });
    expect(manifest.reference.renderer).toBe("word");
    expect(manifest.reference.wordVersion).toBeTruthy();
    expect(manifest.reference.mutoolVersion).toBeTruthy();
  });

  test("retains endpoints for every control under test", async () => {
    const manifest = await readLineEndpointManifest(MANIFEST_PATH);
    const pages = manifest.pages.map(({ lines }) => lines.map(({ text }) => text));

    expect(pages).toHaveLength(20);
    expect(pages[0]?.some((text) => text.endsWith("-"))).toBe(true);
    expect(pages[1]?.some((text) => text.endsWith("-"))).toBe(false);
    expect(pages[2]?.some((text) => text.endsWith("-"))).toBe(false);
    expect(pages[3]?.filter((text) => text.endsWith("-"))).toHaveLength(1);
    expect(pages[4]?.some((text) => text.endsWith("-"))).toBe(true);
    expect(pages[5]?.some((text) => text.endsWith("-"))).toBe(true);
    expect(pages[6]?.slice(1).every((text) => !/^[、。）］]/u.test(text))).toBe(true);
    expect(pages[7]?.slice(1).every((text) => !text.startsWith("※"))).toBe(true);
    expect(pages[8]?.some((text) => text.endsWith("-"))).toBe(true);
    expect(pages[9]?.at(1)).toBe("甲乙丙丁戊己庚辛壬癸子丑");
    expect(pages[10]?.at(1)).toBe("甲乙丙丁戊己庚辛壬癸子丑寅。");
    expect(pages[11]?.at(1)?.endsWith("。")).toBe(true);
    expect(pages[11]?.at(2)?.startsWith("」")).toBe(true);
    expect(pages[12]?.slice(1)).toEqual(pages[10]?.slice(1));
    const commonParagraphLines = [
      "Contract clauses allocate risk between parties and",
      "describe notice requirements, payment timing, remedies,",
      "and termination rights.",
    ];
    expect(pages[13]?.slice(1)).toEqual(commonParagraphLines);
    expect(pages[14]?.slice(1)).toEqual([
      "Contract clauses allocate risk between",
      "parties and describe notice requirements, payment",
      "timing, remedies, and termination rights.",
    ]);
    expect(pages[15]?.at(1)).toBe("Service item Annual renewal and support plan 1,250.00");
    expect(pages[16]?.slice(1)).toEqual([
      "1. The supplier must preserve records, provide audit",
      "access, and notify the customer before deleting",
      "retained material.",
    ]);
    expect(pages[17]?.slice(1)).toEqual([
      "Payment is due within thirty days after",
      "receipt of a valid invoice and supporting",
      "records.",
    ]);
    expect(pages[18]?.slice(1)).toEqual([
      "The agreement requires written notice before",
      "either party may terminate for convenience.",
      "Additional obligations continue.",
    ]);
    expect(pages[19]?.slice(1)).toEqual(commonParagraphLines);
  });
});
