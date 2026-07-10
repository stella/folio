import { describe, expect, test } from "bun:test";

import { parseBlockContent } from "./blockContentParser";
import { parseNumbering } from "./numberingParser";
import { parseParagraph } from "./paragraphParser";
import type { XmlElement } from "./xmlParser";
import { parseXmlDocument } from "./xmlParser";

const NUMBERING_SHARED = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="4">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="(%1)"/>
      <w:pPr><w:ind w:start="0" w:hanging="0"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="4"/></w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="4"/>
    <w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>
  </w:num>
</w:numbering>`;

function parseParagraphXml(xml: string, numbering: ReturnType<typeof parseNumbering>) {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse paragraph XML");
  }
  return parseParagraph(root, null, null, numbering, null, null);
}

describe("paragraphParser exposes abstractNumId and startOverride for sharing", () => {
  const numbering = parseNumbering(NUMBERING_SHARED);

  test("numId without override exposes abstractNumId only", () => {
    const paragraph = parseParagraphXml(
      `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      </w:p>`,
      numbering,
    );

    expect(paragraph.listRendering?.abstractNumId).toBe(4);
    expect(paragraph.listRendering?.startOverride).toBeUndefined();
  });

  test("numId with startOverride exposes the override value", () => {
    const paragraph = parseParagraphXml(
      `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      </w:p>`,
      numbering,
    );

    expect(paragraph.listRendering?.abstractNumId).toBe(4);
    expect(paragraph.listRendering?.startOverride).toBe(1);
  });

  test('level pPr w:start="0" parses to indentLeft=0', () => {
    const level = numbering.getLevel(1, 0);
    expect(level?.pPr?.indentLeft).toBe(0);
  });
});

test("nested numbering starts from authored levels when no parent paragraph exists", () => {
  const numbering = parseNumbering(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="8">
    <w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="3"><w:abstractNumId w:val="8"/></w:num>
</w:numbering>`);
  const root = parseXmlDocument(
    `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
      <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>
    </w:body>`,
  ) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse body XML");
  }

  const paragraphs = parseBlockContent(root, null, null, numbering, null, null);

  expect(paragraphs.at(0)?.type).toBe("paragraph");
  expect(paragraphs.at(0)?.type === "paragraph" && paragraphs.at(0)?.listRendering?.marker).toBe(
    "3.3.",
  );
  expect(paragraphs.at(1)?.type === "paragraph" && paragraphs.at(1)?.listRendering?.marker).toBe(
    "3.4.",
  );
});
