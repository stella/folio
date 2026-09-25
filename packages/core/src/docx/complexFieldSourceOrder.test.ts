/**
 * The runs of a complex field the paragraph does not assemble keep their place.
 *
 * The paragraph walk buffers every run after a `begin` until the `end` that
 * closes it. A `w:hyperlink` or another inline container read in between went
 * straight into the paragraph, ahead of the buffer, and a paragraph that ended
 * with the field still open appended the buffer last. A TOC opened in its first
 * entry's paragraph with `\h` entries therefore saved as the entry's hyperlink
 * followed by the TOC's `begin`, instruction and `separate`: the first entry
 * fell outside the field's result. A field nested in an open field's result (a
 * TOC entry's PAGEREF without `\h`) reset the buffer and discarded the outer
 * field's `begin`, instruction and `separate` altogether.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const run = (inner: string): string => `<w:r>${inner}</w:r>`;
const begin = run('<w:fldChar w:fldCharType="begin"/>');
const separate = run('<w:fldChar w:fldCharType="separate"/>');
const end = run('<w:fldChar w:fldCharType="end"/>');
const instruction = (text: string): string =>
  run(`<w:instrText xml:space="preserve">${text}</w:instrText>`);
const text = (value: string): string => run(`<w:t>${value}</w:t>`);

const TOC_OPEN = `${begin}${instruction(" TOC \\o &quot;1-3&quot; \\h \\z \\u ")}${separate}`;
const PAGE_REF = `${begin}${instruction(" PAGEREF _Toc1 \\h ")}${separate}${text("2")}${end}`;

const savedDocumentXml = async (body: string): Promise<string> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
    preloadFonts: false,
  });
  const saved = await JSZip.loadAsync(await repackDocx(parsed, { updateModifiedDate: false }));
  const xml = (await saved.file("word/document.xml")?.async("text")) ?? "";
  return xml.slice(xml.indexOf("<w:body>") + "<w:body>".length, xml.indexOf("<w:sectPr"));
};

describe("complex field runs keep their source order", () => {
  test("a field closed in a later paragraph stays ahead of the hyperlink its result holds", async () => {
    const firstEntry = `<w:p>${TOC_OPEN}<w:hyperlink w:anchor="_Toc1" w:history="1">${text("Summary")}${run("<w:tab/>")}${PAGE_REF}</w:hyperlink></w:p>`;
    const body = `${firstEntry}<w:p>${text("Second")}</w:p><w:p>${end}</w:p>`;

    expect(await savedDocumentXml(body)).toBe(body);
  });

  test("an outer field whose result holds a nested field keeps its begin, code and separator", async () => {
    const firstEntry = `<w:p>${TOC_OPEN}${text("Summary")}${run("<w:tab/>")}${PAGE_REF}</w:p>`;
    const body = `${firstEntry}<w:p>${text("Second")}</w:p><w:p>${end}</w:p>`;

    expect(await savedDocumentXml(body)).toBe(body);
  });

  test("a field closed in the same paragraph around a hyperlink keeps the hyperlink inside", async () => {
    const body = `<w:p>${TOC_OPEN}<w:hyperlink w:anchor="_Toc1" w:history="1">${text("Summary")}</w:hyperlink>${end}</w:p>`;

    expect(await savedDocumentXml(body)).toBe(body);
  });

  test("markers read inside an unclosed field stay between the same runs", async () => {
    const body = `<w:p>${TOC_OPEN}<w:bookmarkStart w:id="0" w:name="entry"/>${text("Summary")}<w:bookmarkEnd w:id="0"/></w:p><w:p>${end}</w:p>`;

    expect(await savedDocumentXml(body)).toBe(body);
  });

  test("a field closed in its own paragraph is still assembled", async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx());
    zip.file(
      "word/document.xml",
      `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body><w:p>${PAGE_REF}</w:p><w:sectPr/></w:body></w:document>`,
    );
    const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });
    const [paragraph] = parsed.package.document.content;
    expect(paragraph?.type).toBe("paragraph");
    const contents = paragraph?.type === "paragraph" ? paragraph.content : [];
    expect(contents.map((item) => item.type)).toEqual(["complexField"]);
  });
});
