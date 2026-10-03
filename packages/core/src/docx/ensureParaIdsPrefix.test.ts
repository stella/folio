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
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { ensureParaIds, EnsureParaIdsError } from "./ensureParaIds";
import { rewritePackagePrefixes, type PrefixVariant } from "./__tests__/namespacePrefixVariants";
import { createDocx } from "./rezip";
import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  resolveAttributeNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";
import {
  PARAGRAPH_SCAN_NAMES,
  resolveWordprocessingPrefixes,
  splicesAsCanonical,
} from "./wordprocessingPrefixes";

const W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W_URI = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const MC_URI = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PACKAGE_REL_URI = "http://schemas.openxmlformats.org/package/2006/relationships";
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

  test("leaves foreign paragraphs alone when a nested element rebinds the paragraph prefix", async () => {
    const xml =
      `<x:document xmlns:x="${W_URI}"><x:body><x:p><x:r><x:t>a</x:t></x:r></x:p>` +
      `<x:tbl xmlns:x="urn:example:other"><x:p/></x:tbl><x:sectPr/></x:body></x:document>`;
    const input = await withDocumentPart(await twoParagraphs(), xml);
    const result = await ensureParaIds(input);
    expect(result.assigned).toBe(1);
    expect(await documentPart(result.docx)).toContain(
      `<x:tbl xmlns:x="urn:example:other"><x:p/></x:tbl>`,
    );
    expect((await ensureParaIds(result.docx)).alreadyComplete).toBe(true);
  });

  test("never reports a body complete when its paragraphs went unseen", async () => {
    // WordprocessingML bound only on the paragraphs themselves, under a prefix
    // the root never declares: the scan has nothing to match, the parser does.
    const xml =
      `<doc:document xmlns:doc="${W_URI}"><doc:body>` +
      `<y:p xmlns:y="${W_URI}"><y:r><y:t>a</y:t></y:r></y:p><doc:sectPr/></doc:body></doc:document>`;
    const input = await withDocumentPart(await twoParagraphs(), xml);
    const result = await ensureParaIds(input);
    expect(result.assigned).toBe(1);
    expect((await ensureParaIds(result.docx)).alreadyComplete).toBe(true);
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

const namespaceBindings = fc.record({
  namespace: fc.constantFrom(W_URI, STRICT_W_URI),
  prefix: fc.constantFrom("w", "x", ""),
  idPrefix: fc.constantFrom("w14", "ids", "ext"),
  mcPrefix: fc.constantFrom("mc", "compat"),
  extensionBinding: fc.constantFrom("root", "nested", "foreign"),
  mainPart: fc.constantFrom(
    "word/document.xml",
    "word/document2.xml",
    "parts/main.xml",
    "word/my document.xml",
    "parts/článek_日本.xml",
  ),
  relationshipProfile: fc.constantFrom(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
  ),
  relationshipPrefix: fc.constantFrom("", "pkg"),
  targetSpelling: fc.constantFrom("relative", "absolute", "dot", "encoded"),
  paragraphCount: fc.integer({ min: 1, max: 6 }),
  quote: fc.constantFrom('"', "'"),
});

const qualified = (prefix: string, local: string): string =>
  prefix === "" ? local : `${prefix}:${local}`;

const allParagraphs = (root: XmlElement): XmlElement[] => {
  const result: XmlElement[] = [];
  const visit = (element: XmlElement): void => {
    if (
      getLocalName(element.name ?? "") === "p" &&
      WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
    )
      result.push(element);
    for (const child of getChildElements(element)) visit(child);
  };
  visit(root);
  return result;
};

const paragraphIds = (element: XmlElement): string[] =>
  Object.entries(element.attributes ?? {})
    .filter(([name]) => {
      const uri = resolveAttributeNamespaceUri(element, name);
      return (
        getLocalName(name) === "paraId" &&
        (uri === W14_URI || WORDPROCESSINGML_NAMESPACE_URIS.has(uri ?? ""))
      );
    })
    .map(([, value]) => String(value));

test(
  "namespace binding dimensions preserve paragraph coverage, uniqueness and byte idempotence",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        namespaceBindings,
        async ({
          namespace,
          prefix,
          idPrefix,
          mcPrefix,
          extensionBinding,
          mainPart,
          relationshipProfile,
          relationshipPrefix,
          targetSpelling,
          paragraphCount,
          quote,
        }) => {
          const q = (local: string) => qualified(prefix, local);
          const rootDeclaration =
            prefix === ""
              ? `xmlns=${quote}${namespace}${quote}`
              : `xmlns:${prefix}=${quote}${namespace}${quote}`;
          const idDeclaration = `xmlns:${idPrefix}=${quote}${extensionBinding === "foreign" ? "urn:foreign:ids" : W14_URI}${quote}`;
          const rootIdDeclaration = extensionBinding === "nested" ? "" : idDeclaration;
          const nestedIdDeclaration =
            extensionBinding === "nested"
              ? idDeclaration
              : `xmlns:${idPrefix}="urn:foreign:nested"`;
          const existingPrefix = extensionBinding === "nested" ? idPrefix : "old";
          const nestedWordDeclaration =
            prefix === "" ? 'xmlns="urn:foreign:word"' : `xmlns:${prefix}="urn:foreign:word"`;
          const foreign = `<foreign:p xmlns:foreign="urn:foreign" foreign:paraId="FEED0002" data="a&gt;b"/>`;
          const fallback = `<local:Fallback xmlns:local="${MC_URI}"><actual:p xmlns:actual="${namespace}" xmlns:kept="${W14_URI}" kept:paraId="FEED0001"/></local:Fallback>`;
          const repeated = Array.from({ length: paragraphCount }, () => `<${q("p")} />`).join("");
          const xml = `<${q("document")} ${rootDeclaration} ${rootIdDeclaration} xmlns:${mcPrefix}="${MC_URI}" ${mcPrefix}:Ignorable=""><${q("body")}>${repeated}<scope ${nestedWordDeclaration} ${nestedIdDeclaration} xmlns:${mcPrefix}="urn:foreign:compat"><actual:p xmlns:actual="${namespace}" xmlns:old="${W14_URI}" ${existingPrefix}:paraId = ${quote}FEED0001${quote} ${existingPrefix}:textId="00000000"/><actual:p xmlns:actual="${namespace}" xmlns:other="${namespace}"><other:r><other:t>Unchanged</other:t></other:r></actual:p>${foreign}${fallback}</scope><${q("p")}/></${q("body")}></${q("document")}>`;
          const target = {
            relative: mainPart,
            absolute: `/${mainPart}`,
            dot: `./${mainPart}`,
            encoded: mainPart.split("/").map(encodeURIComponent).join("/"),
          }[targetSpelling];
          const rel = qualified(relationshipPrefix, "Relationship");
          const rels = qualified(relationshipPrefix, "Relationships");
          const relDeclaration =
            relationshipPrefix === ""
              ? `xmlns="${PACKAGE_REL_URI}"`
              : `xmlns:${relationshipPrefix}="${PACKAGE_REL_URI}"`;
          const zip = new JSZip();
          zip.file(
            "_rels/.rels",
            `<${rels} ${relDeclaration}><${rel} Id="rId1" Type="${relationshipProfile}/officeDocument" Target="${target}"/></${rels}>`,
          );
          zip.file(mainPart, xml);
          zip.file(
            "word/comments.xml",
            `<w:comments xmlns:w="${W_URI}" xmlns:ids="${W14_URI}"><w:p ids:paraId="FEED0003"/></w:comments>`,
          );
          const input = await zip.generateAsync({ type: "uint8array" });
          const normalized = await ensureParaIds(input);
          expect(normalized.assigned).toBe(paragraphCount + 2);
          expect(normalized.deduplicated).toBe(1);
          const saved = await JSZip.loadAsync(normalized.docx);
          const output = await saved.file(mainPart)?.async("text");
          expect(output).toBeDefined();
          const root = parseXmlDocument(output ?? "");
          if (root === null) throw new Error("normalized fixture has no root");
          const ids = allParagraphs(root).map(paragraphIds);
          expect(ids).toHaveLength(paragraphCount + 4);
          expect(ids.every((values) => values.length === 1)).toBe(true);
          const editable = ids.filter((values) => values.at(0) !== "FEED0001").flat();
          expect(new Set(editable).size).toBe(paragraphCount + 3);
          expect(
            editable.every(
              (id) => /^[0-9A-F]{8}$/u.test(id) && id !== "00000000" && id !== "FEED0003",
            ),
          ).toBe(true);
          expect(output).toContain(foreign);
          expect(output).toContain(fallback);
          expect(output).toContain("<other:t>Unchanged</other:t>");
          expect(await saved.file("word/comments.xml")?.async("text")).toBe(
            await zip.file("word/comments.xml")?.async("text"),
          );
          const again = await ensureParaIds(normalized.docx);
          expect(again.alreadyComplete).toBe(true);
          expect(again.docx).toBe(normalized.docx);
          expect(again.mintedParaIds).toEqual([]);
          const repeat = await ensureParaIds(input);
          expect(await (await JSZip.loadAsync(repeat.docx)).file(mainPart)?.async("text")).toBe(
            output,
          );
        },
      ),
      { numRuns: 60 },
    );
  },
  propertyTestTimeout(30_000),
);

for (const malformed of [
  `<w:document xmlns:w="${W_URI}"><w:body></w:document>`,
  `<w:document xmlns:w="${W_URI}"><w:p key="unterminated/></w:document>`,
]) {
  test("refuses malformed paragraph markup", async () => {
    await expect(
      ensureParaIds(await withDocumentPart(await twoParagraphs(), malformed)),
    ).rejects.toThrow(EnsureParaIdsError);
  });
}

test("a missing officeDocument target is refused at the attribute boundary", async () => {
  const zip = new JSZip();
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="${PACKAGE_REL_URI}"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W_URI}"><w:body><w:p/></w:body></w:document>`,
  );
  await expect(
    ensureParaIds(await zip.generateAsync({ type: "uint8array" })),
  ).rejects.toBeInstanceOf(EnsureParaIdsError);
});
