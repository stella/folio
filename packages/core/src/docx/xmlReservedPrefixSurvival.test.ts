/**
 * The `xml:` prefix survives a save without being declared.
 *
 * The XML Namespaces recommendation reserves `xml` as bound to
 * `http://www.w3.org/XML/1998/namespace` in every document, whether or not
 * the document writes `xmlns:xml`, and almost none do. A container's
 * attribute remainder resolved a prefix's namespace URI to decide whether it
 * could spell the attribute back; with no declaration in scope, `xml:space`
 * resolved to no namespace at all and was written back unprefixed —
 * `<w:p space="preserve">` — which is not the XML namespace attribute the
 * source wrote and is schema-invalid. The same loss reached every container
 * whose attributes round-trip through the remainder: `w:p`, `w:r`, `w:tr`,
 * `w:sectPr`.
 *
 * `w:t`, `w:instrText` and `w:delText` are absent on purpose: their parser
 * never reads `xml:space` into the remainder because the serializer
 * re-derives it from the text content (`requiresXmlSpacePreserve`), so this
 * loss never reached them.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}">` +
      `<w:body>${body}<w:sectPr xml:lang="en-US"/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentXmlOf = async (saved: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";

/** Parse and save, with no editor round trip in between: the plainest resave. */
const roundTrip = async (body: string): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  return documentXmlOf(await repackDocx(parsed, { updateModifiedDate: false }));
};

describe("xml: attributes on unmodelled containers", () => {
  test("w:p keeps xml:space prefixed", async () => {
    const xml = await roundTrip('<w:p xml:space="preserve"><w:r><w:t>a</w:t></w:r></w:p>');
    expect(xml).toContain('xml:space="preserve"');
    expect(xml).not.toMatch(/<w:p[^>]*\sspace="preserve"/);
  });

  test("w:r keeps xml:space prefixed", async () => {
    const xml = await roundTrip('<w:p><w:r xml:space="preserve"><w:t> a</w:t></w:r></w:p>');
    expect(xml).toContain('xml:space="preserve"');
    expect(xml).not.toMatch(/<w:r[^>]*\sspace="preserve"/);
  });

  test("w:tr keeps xml:space prefixed", async () => {
    const xml = await roundTrip(
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
        '<w:tr xml:space="preserve"><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    );
    expect(xml).toContain('xml:space="preserve"');
    expect(xml).not.toMatch(/<w:tr[^>]*\sspace="preserve"/);
  });

  test("w:sectPr keeps xml:lang prefixed", async () => {
    const xml = await roundTrip("<w:p><w:r><w:t>a</w:t></w:r></w:p>");
    expect(xml).toContain('xml:lang="en-US"');
    expect(xml).not.toMatch(/<w:sectPr[^>]*\slang="en-US"/);
  });
});
