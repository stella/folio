import { describe, expect, test } from "bun:test";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const paragraphFromXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml);
  if (!root) throw new Error("Failed to parse the paragraph fixture");
  return parseParagraph(root, null, null, null);
};

const roundTrip = (source: Document): Document =>
  fromProseDoc(toProseDoc(source), source, {
    stylesheetSource: { type: "package" },
  });

describe("comment range boundaries in inline content controls", () => {
  test("parsing and saving keeps direct anchors inside sdtContent", () => {
    const paragraph = paragraphFromXml(
      `<w:p xmlns:w="${W}"><w:commentRangeStart w:id="9"/>` +
        `<w:sdt><w:sdtPr><w:tag w:val="clause"/></w:sdtPr>` +
        `<w:sdtContent><w:r><w:t>covered</w:t></w:r>` +
        `<w:commentRangeEnd w:id="9"/></w:sdtContent></w:sdt>` +
        `<w:r><w:commentReference w:id="9"/></w:r></w:p>`,
    );
    const control = paragraph.content.find((item) => item.type === "inlineSdt");

    expect(control).toMatchObject({
      type: "inlineSdt",
      content: [
        { type: "run", content: [{ type: "text", text: "covered" }] },
        { type: "commentRangeEnd", id: 9 },
      ],
    });

    const saved = serializeParagraph(paragraph);
    const sdtContentStart = saved.indexOf("<w:sdtContent>");
    const sdtContentEnd = saved.indexOf("</w:sdtContent>", sdtContentStart);
    expect(sdtContentStart).toBeGreaterThanOrEqual(0);
    expect(saved.slice(sdtContentStart, sdtContentEnd)).toContain(
      '<w:commentRangeEnd w:id="9"/>',
    );

    const reparsed = paragraphFromXml(saved);
    expect(reparsed.content).toEqual(paragraph.content);
  });

  test("a nonempty range inside the control closes before outside text", () => {
    const paragraph = paragraphFromXml(
      `<w:p xmlns:w="${W}"><w:sdt><w:sdtPr>` +
        `<w:tag w:val="clause"/></w:sdtPr><w:sdtContent>` +
        `<w:commentRangeStart w:id="9"/><w:r><w:t>inside</w:t></w:r>` +
        `<w:commentRangeEnd w:id="9"/></w:sdtContent></w:sdt>` +
        `<w:r><w:t> outside</w:t></w:r>` +
        `<w:r><w:commentReference w:id="9"/></w:r></w:p>`,
    );
    const template = createEmptyDocument();
    const source: Document = {
      ...template,
      package: {
        ...template.package,
        document: {
          ...template.package.document,
          content: [paragraph],
          comments: [
            {
              id: 9,
              author: "Reviewer",
              date: "2026-01-01T00:00:00Z",
              content: [],
            },
          ],
        },
      },
    };

    const editor = toProseDoc(source);
    let insideHasComment = false;
    let outsideHasComment = false;
    editor.descendants((node) => {
      if (!node.isText) return true;
      if (node.text?.includes("inside")) {
        insideHasComment = node.marks.some(
          (mark) => mark.type.name === "comment",
        );
      }
      if (node.text?.includes("outside")) {
        outsideHasComment = node.marks.some(
          (mark) => mark.type.name === "comment",
        );
      }
      return true;
    });
    expect(insideHasComment).toBe(true);
    expect(outsideHasComment).toBe(false);
    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });
});
