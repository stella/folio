import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { extractDocxText } from "./extractDocxText";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type MakeDocxOptions = {
  body: string;
  headers?: Record<string, string>;
  footers?: Record<string, string>;
};

const makeDocx = async ({
  body,
  headers = {},
  footers = {},
}: MakeDocxOptions): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`,
  );
  for (const [path, content] of Object.entries(headers)) {
    zip.file(path, `<w:hdr xmlns:w="${W_NS}">${content}</w:hdr>`);
  }
  for (const [path, content] of Object.entries(footers)) {
    zip.file(path, `<w:ftr xmlns:w="${W_NS}">${content}</w:ftr>`);
  }
  return await zip.generateAsync({ type: "uint8array" });
};

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe("extractDocxText", () => {
  test("returns deterministic paragraphs across document parts", async () => {
    const bytes = await makeDocx({
      body:
        paragraph("Before") +
        `<w:tbl><w:tr><w:tc>${paragraph("Cell")}</w:tc></w:tr></w:tbl>` +
        `<w:sdt><w:sdtContent>${paragraph("Control")}</w:sdtContent></w:sdt>` +
        paragraph("After"),
      headers: {
        "word/header2.xml": paragraph("Header 2"),
        "word/header1.xml": paragraph("Header 1"),
      },
      footers: {
        "word/footer1.xml": paragraph("Footer"),
      },
    });

    const result = await extractDocxText(bytes);

    expect(
      result.paragraphs.map(({ index, text, source }) => ({
        index,
        text,
        source,
      })),
    ).toEqual([
      { index: 0, text: "Header 1", source: "header" },
      { index: 1, text: "Header 2", source: "header" },
      { index: 2, text: "Before", source: "body" },
      { index: 3, text: "Cell", source: "body" },
      { index: 4, text: "Control", source: "body" },
      { index: 5, text: "After", source: "body" },
      { index: 6, text: "Footer", source: "footer" },
    ]);
    expect(result.charCount).toBe("Header 1Header 2BeforeCellControlAfterFooter".length);
    expect(result.view).toBe("accepted");
  });

  test("preserves explicit whitespace and applies the accepted revision view", async () => {
    const bytes = await makeDocx({
      body: `<w:p>
        <w:r><w:t>Keep</w:t><w:br/><w:t>Line</w:t><w:tab/><w:t>Tab</w:t></w:r>
        <w:del><w:r><w:delText>Deleted</w:delText></w:r></w:del>
        <w:moveFrom><w:r><w:t>Moved away</w:t></w:r></w:moveFrom>
        <w:ins><w:r><w:t>Inserted</w:t></w:r></w:ins>
      </w:p>`,
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.at(0)?.text).toBe("Keep\nLine\tTabInserted");
  });

  test("extracts paragraph and majority-run formatting metadata", async () => {
    const bytes = await makeDocx({
      body: `<w:p>
        <w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Majority</w:t></w:r>
        <w:r><w:t>x</w:t></w:r>
      </w:p>
      <w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Left</w:t></w:r></w:p>`,
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.at(0)).toEqual({
      index: 0,
      text: "Majorityx",
      source: "body",
      style: "Heading1",
      alignment: "center",
      bold: true,
      fontSize: 28,
    });
    expect(result.paragraphs.at(1)?.alignment).toBe("left");
  });

  test("supports alternate prefixes for the main OOXML namespace", async () => {
    const zip = new JSZip();
    zip.file(
      "word/document.xml",
      `<x:document xmlns:x="${W_NS}"><x:body><x:p><x:pPr><x:pStyle x:val="Title"/></x:pPr><x:r><x:t>Alternate</x:t></x:r></x:p></x:body></x:document>`,
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.at(0)?.text).toBe("Alternate");
    expect(result.paragraphs.at(0)?.style).toBe("Title");
  });

  test("returns an empty result when the main document part is absent", async () => {
    const zip = new JSZip();
    zip.file("other.xml", "<root/>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    expect(await extractDocxText(bytes)).toEqual({
      paragraphs: [],
      charCount: 0,
      view: "accepted",
    });
  });
});
