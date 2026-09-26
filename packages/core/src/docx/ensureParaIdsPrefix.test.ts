/**
 * `ensureParaIds` reads WordprocessingML by namespace, not by the `w:` prefix.
 *
 * A package that binds WordprocessingML to `x:` (or to the default namespace)
 * opens like any other, but the normalizer matched `<w:p` literally, saw no
 * paragraphs, and reported `alreadyComplete: true` with nothing stamped: every
 * block stayed positional while the caller was told its ids were stable.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { ensureParaIds, EnsureParaIdsError } from "./ensureParaIds";
import { rewritePackagePrefixes, type PrefixVariant } from "./__tests__/namespacePrefixVariants";
import { createDocx } from "./rezip";
import {
  PARAGRAPH_SCAN_NAMES,
  resolveWordprocessingPrefixes,
  splicesAsCanonical,
} from "./wordprocessingPrefixes";

const W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_URI = "http://schemas.microsoft.com/office/word/2010/wordml";

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

const twoParagraphs = async (): Promise<Uint8Array> =>
  new Uint8Array(await createDocx(fromMarkdown("Alpha clause.\n\nBeta clause.")));

const documentPart = async (docx: Uint8Array): Promise<string> => {
  const text = await (await JSZip.loadAsync(docx)).file("word/document.xml")?.async("text");
  if (text === undefined) throw new Error("word/document.xml missing");
  return text;
};

const withDocumentPart = async (docx: Uint8Array, xml: string): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(docx);
  zip.file("word/document.xml", xml);
  return zip.generateAsync({ type: "uint8array" });
};

const respelled = async (docx: Uint8Array, variant: PrefixVariant): Promise<Uint8Array> => {
  const result = await rewritePackagePrefixes(docx, variant);
  if (!result) throw new Error(`fixture cannot be respelled as ${variant}`);
  return result.docx;
};

describe("ensureParaIds with non-conventional WordprocessingML prefixes", () => {
  for (const variant of ["alias", "default", "w14-alias"] as const) {
    test(`${variant}: stamps every paragraph and the ids survive save and reopen`, async () => {
      const input = await respelled(await twoParagraphs(), variant);
      const before = await FolioDocxReviewer.fromBuffer(toArrayBuffer(input));
      expect(before.getContent().map((block) => block.text)).toEqual([
        "Alpha clause.",
        "Beta clause.",
      ]);

      const result = await ensureParaIds(input);
      expect(result.alreadyComplete).toBe(false);
      expect(result.assigned).toBe(2);

      const reviewer = await FolioDocxReviewer.fromBuffer(toArrayBuffer(result.docx));
      const blocks = reviewer.getContent();
      expect(blocks.map((block) => block.idStability)).toEqual([undefined, undefined]);

      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      expect(reopened.getContent().map((block) => block.id)).toEqual(
        blocks.map((block) => block.id),
      );

      const again = await ensureParaIds(result.docx);
      expect(again.alreadyComplete).toBe(true);
      expect(again.assigned).toBe(0);
    });
  }

  test("writes the ids under the part's own w14 prefix and lists that prefix as ignorable", async () => {
    const xml =
      `<x:document xmlns:x="${W_URI}" xmlns:x14="${W14_URI}" ` +
      `xmlns:m2="http://schemas.openxmlformats.org/markup-compatibility/2006" m2:Ignorable="x14">` +
      `<x:body><x:p x14:paraId="1A2B3C4D"><x:r><x:t>Kept</x:t></x:r></x:p>` +
      `<x:p><x:r><x:t>Stamped</x:t></x:r></x:p><x:sectPr/></x:body></x:document>`;
    const result = await ensureParaIds(await withDocumentPart(await twoParagraphs(), xml));
    expect(result.assigned).toBe(1);

    const out = await documentPart(result.docx);
    expect(out).toContain('x14:paraId="1A2B3C4D"');
    expect(out.match(/\sx14:paraId=/gu)).toHaveLength(2);
    expect(out).not.toContain("w14:");
    expect(out).toContain('m2:Ignorable="x14"');
    expect(out).not.toContain("xmlns:w14");
  });

  test("leaves mc:Fallback paragraphs alone under an alternate mc prefix", async () => {
    const xml =
      `<x:document xmlns:x="${W_URI}" ` +
      `xmlns:m2="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
      `<x:body><x:p><x:r><m2:AlternateContent><m2:Choice Requires="x"><x:t>c</x:t></m2:Choice>` +
      `<m2:Fallback><x:p><x:r><x:t>f</x:t></x:r></x:p></m2:Fallback></m2:AlternateContent></x:r></x:p>` +
      `<x:sectPr/></x:body></x:document>`;
    const result = await ensureParaIds(await withDocumentPart(await twoParagraphs(), xml));
    expect(result.assigned).toBe(1);
    const out = await documentPart(result.docx);
    expect(out).toContain("<m2:Fallback><x:p><x:r>");
    expect(out).toContain('m2:Ignorable="w14"');
  });

  test("refuses a part whose nested element rebinds the paragraph prefix", async () => {
    const xml =
      `<x:document xmlns:x="${W_URI}"><x:body><x:p><x:r><x:t>a</x:t></x:r></x:p>` +
      `<x:tbl xmlns:x="urn:example:other"><x:p/></x:tbl><x:sectPr/></x:body></x:document>`;
    const input = await withDocumentPart(await twoParagraphs(), xml);
    await expect(ensureParaIds(input)).rejects.toThrow(EnsureParaIdsError);
    await expect(ensureParaIds(input)).rejects.toThrow("rebinds prefix x");
  });

  test("never reports a body complete when its paragraphs went unseen", async () => {
    // WordprocessingML bound only on the paragraphs themselves, under a prefix
    // the root never declares: the scan has nothing to match, the parser does.
    const xml =
      `<doc:document xmlns:doc="${W_URI}"><doc:body>` +
      `<y:p xmlns:y="${W_URI}"><y:r><y:t>a</y:t></y:r></y:p><doc:sectPr/></doc:body></doc:document>`;
    const input = await withDocumentPart(await twoParagraphs(), xml);
    await expect(ensureParaIds(input)).rejects.toThrow(EnsureParaIdsError);
  });
});

describe("resolveWordprocessingPrefixes", () => {
  test("reads the conventional spelling as canonical", () => {
    const resolution = resolveWordprocessingPrefixes(
      `<w:document xmlns:w="${W_URI}" xmlns:w14="${W14_URI}"><w:body/></w:document>`,
    );
    expect(resolution).toMatchObject({
      type: "resolved",
      prefixes: { main: ["w"], w14: ["w14"], canonical: true },
    });
  });

  test("resolves a default namespace and a second WordprocessingML prefix", () => {
    const resolution = resolveWordprocessingPrefixes(
      `<document xmlns="${W_URI}" xmlns:x="${W_URI}"><body/></document>`,
    );
    expect(resolution).toMatchObject({
      type: "resolved",
      prefixes: { main: ["", "x"], canonical: false },
    });
  });

  test("accepts a nested declaration that repeats a binding", () => {
    const resolution = resolveWordprocessingPrefixes(
      `<w:document xmlns:w="${W_URI}"><w:body><w:p xmlns:w="${W_URI}"/></w:body></w:document>`,
    );
    expect(resolution).toMatchObject({ type: "resolved", prefixes: { canonical: true } });
  });

  test("reports a nested WordprocessingML binding under a new prefix as unsupported", () => {
    const resolution = resolveWordprocessingPrefixes(
      `<w:document xmlns:w="${W_URI}"><w:body><y:p xmlns:y="${W_URI}"/></w:body></w:document>`,
    );
    expect(resolution.type).toBe("unsupported");
  });

  test("a second WordprocessingML alias blocks a paragraph splice only where it spells a paragraph", () => {
    const root = `<w:document xmlns:w="${W_URI}" xmlns:x="${W_URI}"><w:body>`;
    expect(
      splicesAsCanonical(`${root}<w:p><x:date x:val="d"/></w:p></w:body></w:document>`, [
        ...PARAGRAPH_SCAN_NAMES,
      ]),
    ).toBe(true);
    expect(
      splicesAsCanonical(`${root}<x:p/></w:body></w:document>`, [...PARAGRAPH_SCAN_NAMES]),
    ).toBe(false);
  });

  test("ignores declarations inside comments and CDATA", () => {
    const resolution = resolveWordprocessingPrefixes(
      `<w:document xmlns:w="${W_URI}"><!-- <w:p xmlns:w="urn:x"/> --><w:body/></w:document>`,
    );
    expect(resolution).toMatchObject({ type: "resolved", prefixes: { canonical: true } });
  });
});
