import { describe, expect, test } from "bun:test";

import { parseRunProperties } from "../../docx/runParser";
import { parseXml } from "../../docx/xmlParser";
import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { schema } from "../../prosemirror/schema";
import type { Document, TextFormatting } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

const parseFormatting = (innerXml: string): TextFormatting => {
  const parsed = parseXml(
    `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${innerXml}</w:rPr>`,
  );
  const runProperties = parsed.elements.at(0);
  if (!runProperties || runProperties.type !== "element") {
    throw new TypeError("Expected run properties element");
  }
  const formatting = parseRunProperties(runProperties, null);
  if (!formatting) {
    throw new TypeError("Expected parsed run formatting");
  }
  return formatting;
};

const documentWithRun = (text: string, formatting: TextFormatting): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              formatting,
              content: [{ type: "text", text }],
            },
          ],
        },
      ],
    },
  },
});

const firstDocumentRunFormatting = (document: Document): TextFormatting | undefined => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new TypeError("Expected paragraph");
  }
  const run = paragraph.content.at(0);
  if (run?.type !== "run") {
    throw new TypeError("Expected run");
  }
  return run.formatting;
};

const firstFlowTextRun = (document: Document) => {
  const paragraph = toFlowBlocks(toProseDoc(document), {}).find(
    (block) => block.kind === "paragraph",
  );
  if (!paragraph) {
    throw new TypeError("Expected paragraph block");
  }
  const run = paragraph.runs.find((candidate) => candidate.kind === "text");
  if (!run || run.kind !== "text") {
    throw new TypeError("Expected text run");
  }
  return run;
};

describe("complex-script formatting pipeline", () => {
  test("keeps independent negative and positive overrides from OOXML through Flow", () => {
    const formatting = parseFormatting(
      '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Traditional Arabic"/>' +
        '<w:b/><w:bCs w:val="0"/><w:i w:val="0"/><w:iCs/>' +
        '<w:sz w:val="20"/><w:szCs w:val="32"/>',
    );
    const document = documentWithRun("Aع", formatting);

    const roundTripped = firstDocumentRunFormatting(fromProseDoc(toProseDoc(document)));
    expect(roundTripped).toMatchObject({
      bold: true,
      boldCs: false,
      italic: false,
      italicCs: true,
      fontSize: 20,
      fontSizeCs: 32,
      fontFamily: { ascii: "Arial", hAnsi: "Arial", cs: "Traditional Arabic" },
    });

    const flowRun = firstFlowTextRun(document);
    expect(flowRun).toMatchObject({
      fontFamily: "Arial",
      complexScriptFontFamily: "Traditional Arabic",
      fontSize: 10,
      complexScriptFontSize: 16,
      bold: true,
      complexScriptBold: false,
      italic: false,
      complexScriptItalic: true,
    });
  });

  test("keeps combined ordinary and complex-script negatives through a JSON clone", () => {
    const document = documentWithRun("Aع", {
      bold: false,
      boldCs: false,
      italic: false,
      italicCs: false,
    });
    const proseDoc = toProseDoc(document);
    const clonedProseDoc = schema.nodeFromJSON(proseDoc.toJSON());

    expect(firstDocumentRunFormatting(fromProseDoc(clonedProseDoc))).toEqual({
      bold: false,
      boldCs: false,
      italic: false,
      italicCs: false,
    });

    const paragraph = toFlowBlocks(clonedProseDoc, {}).find((block) => block.kind === "paragraph");
    const run = paragraph?.runs.find((candidate) => candidate.kind === "text");
    expect(run).toMatchObject({
      bold: false,
      complexScriptBold: false,
      complexScriptItalic: false,
      italic: false,
    });
  });

  test("force-CS selects complex-script formatting even for Latin text", () => {
    const formatting = parseFormatting(
      '<w:rFonts w:ascii="Arial" w:cs="Traditional Arabic"/>' +
        '<w:b/><w:bCs w:val="0"/><w:iCs/><w:sz w:val="20"/><w:szCs w:val="30"/><w:cs/>',
    );
    const flowRun = firstFlowTextRun(documentWithRun("Latin", formatting));

    expect(flowRun.forceComplexScript).toBe(true);
    expect(flowRun).toMatchObject({
      complexScriptFontFamily: "Traditional Arabic",
      complexScriptFontSize: 15,
      complexScriptBold: false,
      complexScriptItalic: true,
    });
  });

  test.each([
    ["bold", [schema.mark("bold")]],
    ["italic", [schema.mark("italic")]],
    ["italic and strike", [schema.mark("italic"), schema.mark("strike")]],
    ["font size", [schema.mark("fontSize", { size: 24 })]],
  ])("does not synthesize a CS override for ordinary %s marks", (_label, marks) => {
    const source = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("!", marks)]),
    ]);
    const roundTripped = toProseDoc(fromProseDoc(source));
    const text = roundTripped.firstChild?.firstChild;

    expect(text?.marks.some((mark) => mark.type.name === "runFormattingOverride")).toBe(false);
  });
});
