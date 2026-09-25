import { describe, expect, test } from "bun:test";

import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph } from "../types/document";
import { parseBlockContent } from "./blockContentParser";
import { convertBulletToUnicode } from "./bulletMarkers";
import { parseNumbering } from "./numberingParser";
import type { XmlElement } from "./xmlParser";
import { parseXmlDocument } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const bulletLevel = (ilvl: number, text: string, font?: string): string =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
  `<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
  (font === undefined
    ? ""
    : `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>`) +
  `</w:lvl>`;

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W}">
  <w:abstractNum w:abstractNumId="0">
    ${bulletLevel(0, "o", "Courier New")}
    ${bulletLevel(1, "o", "Wingdings")}
    ${bulletLevel(2, "o")}
    ${bulletLevel(3, "§", "Arial")}
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

const parseListParagraphs = (levels: number[]): Paragraph[] => {
  const numbering = parseNumbering(NUMBERING);
  const body = levels
    .map(
      (ilvl) =>
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
        `<w:r><w:t>Item</w:t></w:r></w:p>`,
    )
    .join("");
  const root = parseXmlDocument(`<w:body xmlns:w="${W}">${body}</w:body>`) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse body XML");
  }
  return parseBlockContent(root, null, null, numbering, null, null).filter(
    (block): block is Paragraph => block.type === "paragraph",
  );
};

const layoutMarkers = (paragraphs: Paragraph[]): (string | undefined)[] => {
  const document: Document = { package: { document: { content: paragraphs } } };
  return toFlowBlocks(toProseDoc(document))
    .filter((block) => block.kind === "paragraph")
    .map((block) => block.attrs?.listMarker);
};

describe("bullet w:lvlText characters", () => {
  test("paint a Latin character in a text font as the character itself", () => {
    const paragraphs = parseListParagraphs([0, 3]);

    expect(paragraphs.map((paragraph) => paragraph.listRendering?.marker)).toEqual(["o", "§"]);
    expect(layoutMarkers(paragraphs)).toEqual(["o", "§"]);
  });

  test("read the same character as a symbol-font code in a symbol face or without a face", () => {
    const paragraphs = parseListParagraphs([1, 2]);

    expect(paragraphs.map((paragraph) => paragraph.listRendering?.marker)).toEqual(["○", "○"]);
    expect(layoutMarkers(paragraphs)).toEqual(["○", "○"]);
  });

  test("keeps symbol-font private-use codes mapped whatever face is named", () => {
    expect(convertBulletToUnicode("", "Courier New")).toBe("•");
    expect(convertBulletToUnicode("", "Wingdings")).toBe("■");
  });
});
