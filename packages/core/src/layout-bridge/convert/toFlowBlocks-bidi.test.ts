import { describe, expect, test } from "bun:test";

import { parseParagraph } from "../../docx/paragraphParser";
import { parseXmlDocument } from "../../docx/xmlParser";
import { expectParagraphAttrs } from "../../prosemirror/attrs";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

const bidiElement = (bidi?: boolean): string => {
  if (bidi === undefined) {
    return "";
  }
  return bidi ? "<w:bidi/>" : '<w:bidi w:val="0"/>';
};

const documentWithParagraph = (text: string, bidi?: boolean): Document => {
  const root = parseXmlDocument(`
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:pPr>${bidiElement(bidi)}</w:pPr>
      <w:r><w:t>${text}</w:t></w:r>
    </w:p>
  `);
  if (!root) {
    throw new Error("Failed to parse paragraph XML fixture");
  }
  const paragraph = parseParagraph(root, null, null, null, null, null);

  return { package: { document: { content: [paragraph] } } };
};

const projectBidi = (text: string, bidi?: boolean) => {
  const pmDoc = toProseDoc(documentWithParagraph(text, bidi));
  const pmAttrs = expectParagraphAttrs(pmDoc.child(0));
  const block = toFlowBlocks(pmDoc).at(0);
  if (block?.kind !== "paragraph") {
    throw new Error("Expected paragraph flow block");
  }

  return { direction: pmAttrs.direction, bidi: block.attrs?.bidi };
};

describe("toFlowBlocks paragraph bidi tri-state", () => {
  test("preserves explicit LTR for mixed Arabic and Latin text", () => {
    expect(projectBidi("اتفاقية Folio 2026", false)).toEqual({
      direction: { source: "manual", value: "ltr" },
      bidi: false,
    });
  });

  test("preserves explicit RTL", () => {
    expect(projectBidi("اتفاقية", true)).toEqual({
      direction: { source: "manual", value: "rtl" },
      bidi: true,
    });
  });

  test("leaves an undecided paragraph absent for painter auto-direction", () => {
    expect(projectBidi("اتفاقية")).toEqual({ direction: undefined, bidi: undefined });
  });
});
