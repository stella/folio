import { expect, test } from "bun:test";

import { elementToXml, parseXml } from "@stll/folio-core/docx/xmlParser";

import { childElements, type XmlElement } from "./ooxml-package";
import { paragraphText, replaceParagraphText } from "./paragraph-text";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const paragraphOf = (inner: string): XmlElement => {
  const root = parseXml(`<w:p xmlns:w="${W}">${inner}</w:p>`);
  const paragraph = childElements(root).at(0);
  if (paragraph === undefined) throw new Error("no paragraph");
  return paragraph;
};

test("replaces text across runs, keeping the first run's formatting and skipping deletions", () => {
  const paragraph = paragraphOf(
    '<w:r><w:t xml:space="preserve">Dear {{na</w:t></w:r>' +
      "<w:r><w:rPr><w:b/></w:rPr><w:t>me}}, see </w:t></w:r>" +
      "<w:del><w:r><w:delText>old </w:delText></w:r></w:del>" +
      "<w:r><w:t>{{ref}}.</w:t></w:r>",
  );
  expect(paragraphText(paragraph)).toBe("Dear {{name}}, see {{ref}}.");
  const applied = replaceParagraphText(paragraph, [
    { start: 5, end: 13, text: "Jane" },
    { start: 19, end: 26, text: "A-1" },
    // Overlaps the first and is ignored.
    { start: 6, end: 8, text: "x" },
  ]);
  expect(applied).toBe(2);
  expect(paragraphText(paragraph)).toBe("Dear Jane, see A-1.");
  const xml = elementToXml(paragraph);
  expect(xml).toContain('<w:t xml:space="preserve">Dear Jane</w:t>');
  expect(xml).toContain('<w:b/></w:rPr><w:t xml:space="preserve">, see </w:t>');
  expect(xml).toContain("<w:delText>old </w:delText>");
});
