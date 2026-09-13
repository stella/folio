import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { compareDocx } from "./compare";
import { createDocx } from "../docx/rezip";
import type { StyleDefinitions, TextFormatting, Theme } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { importReferencedStyleDefinitions } from "./style-resources";

const destinationStyles: StyleDefinitions = {
  docDefaults: { rPr: { fontSize: 22 } },
  styles: [{ styleId: "Normal", type: "paragraph", default: true }],
};

const theme: Theme = { colorScheme: { accent1: "4472C4" } };

const documentWith = async ({
  styles,
  text,
  formatting,
}: {
  styles: StyleDefinitions;
  text: string;
  formatting?: TextFormatting;
}): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.styles = styles;
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "10000001",
      content: [
        {
          type: "run",
          ...(formatting !== undefined && { formatting }),
          content: [{ type: "text", text }],
        },
      ],
    },
  ];
  return createDocx(document);
};

const stylesXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/styles.xml");
  if (!file) throw new Error("Expected styles.xml");
  return file.async("text");
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("Expected document.xml");
  return file.async("text");
};

describe("referenced comparison style resources", () => {
  test("imports the transitive style closure without changing existing definitions", () => {
    const sourceStyles: StyleDefinitions = {
      ...destinationStyles,
      styles: [
        ...destinationStyles.styles,
        { styleId: "BaseCharacter", type: "character", rPr: { underline: { style: "single" } } },
        {
          styleId: "LinkedParagraph",
          type: "paragraph",
          basedOn: "Normal",
          link: "LinkedCharacter",
        },
        {
          styleId: "LinkedCharacter",
          type: "character",
          basedOn: "BaseCharacter",
          link: "LinkedParagraph",
        },
      ],
    };

    const result = importReferencedStyleDefinitions({
      sourceStyles,
      destinationStyles,
      sourceTheme: theme,
      destinationTheme: theme,
      referencedStyleIds: ["LinkedCharacter"],
    });

    expect(result.status).toBe("imported");
    if (result.status !== "imported") return;
    expect(result.importedStyleIds).toEqual([
      "BaseCharacter",
      "LinkedParagraph",
      "LinkedCharacter",
    ]);
    expect(result.styles?.styles).toEqual(sourceStyles.styles);
    expect(result.styleIdMap.get("LinkedCharacter")).toBe("LinkedCharacter");
    expect(destinationStyles.styles).toEqual([
      { styleId: "Normal", type: "paragraph", default: true },
    ]);
  });

  test("materializes an imported style's effective formatting when document defaults differ", () => {
    const sourceStyles: StyleDefinitions = {
      docDefaults: { rPr: { fontSize: 24 } },
      styles: [
        { styleId: "Normal", type: "paragraph", default: true },
        { styleId: "TargetCharacter", type: "character" },
      ],
    };

    const result = importReferencedStyleDefinitions({
      sourceStyles,
      destinationStyles,
      sourceTheme: theme,
      destinationTheme: theme,
      referencedStyleIds: ["TargetCharacter"],
    });

    expect(result.status).toBe("imported");
    if (result.status !== "imported") return;
    expect(result.styles?.styles.at(-1)).toEqual({
      styleId: "TargetCharacter",
      type: "character",
      rPr: { fontSize: 24 },
    });
  });

  test("materializes theme font references before importing across themes", () => {
    const sourceTheme: Theme = { fontScheme: { minorFont: { latin: "Source Font" } } };
    const result = importReferencedStyleDefinitions({
      sourceStyles: {
        ...destinationStyles,
        styles: [
          ...destinationStyles.styles,
          {
            styleId: "TargetCharacter",
            type: "character",
            rPr: { fontFamily: { ascii: "Source Font", asciiTheme: "minorAscii" } },
          },
        ],
      },
      destinationStyles,
      sourceTheme,
      destinationTheme: { fontScheme: { minorFont: { latin: "Destination Font" } } },
      referencedStyleIds: ["TargetCharacter"],
    });

    expect(result.status).toBe("imported");
    if (result.status !== "imported") return;
    expect(result.styles?.styles.at(-1)?.rPr?.fontFamily).toEqual({ ascii: "Source Font" });
  });

  test("materializes a locale-specific complex-script theme font", () => {
    const result = importReferencedStyleDefinitions({
      sourceStyles: {
        ...destinationStyles,
        styles: [
          ...destinationStyles.styles,
          {
            styleId: "TargetCharacter",
            type: "character",
            rPr: {
              fontFamily: { csTheme: "minorBidi" },
              language: { bidi: "ar-SA" },
            },
          },
        ],
      },
      destinationStyles,
      sourceTheme: {
        fontScheme: {
          minorFont: { latin: "Fallback Font", cs: "", fonts: { Arab: "Arabic Font" } },
        },
      },
      destinationTheme: { fontScheme: { minorFont: { latin: "Destination Font" } } },
      referencedStyleIds: ["TargetCharacter"],
    });

    expect(result.status).toBe("imported");
    if (result.status !== "imported") return;
    expect(result.styles?.styles.at(-1)?.rPr?.fontFamily).toEqual({ cs: "Arabic Font" });
  });

  test("isolates a target style that reuses a conflicting base identifier", () => {
    const result = importReferencedStyleDefinitions({
      sourceStyles: {
        ...destinationStyles,
        styles: [{ styleId: "Normal", type: "paragraph", default: true, rPr: { bold: true } }],
      },
      destinationStyles,
      sourceTheme: theme,
      destinationTheme: theme,
      referencedStyleIds: ["Normal"],
    });

    expect(result.status).toBe("imported");
    if (result.status !== "imported") return;
    expect(result.styleIdMap.get("Normal")).toBe("FolioImportedStyle1");
    expect(result.styles?.styles.at(-1)).toEqual({
      styleId: "FolioImportedStyle1",
      type: "paragraph",
      rPr: { bold: true },
    });
  });

  test("materializes omitted default spacing over conflicting base defaults", () => {
    const result = importReferencedStyleDefinitions({
      sourceStyles: {
        docDefaults: { rPr: { fontSize: 22 } },
        styles: [{ styleId: "Normal", type: "paragraph", default: true }],
      },
      destinationStyles: {
        docDefaults: {
          pPr: { spaceBefore: 80, spaceAfter: 200, lineSpacing: 276, lineSpacingRule: "auto" },
          rPr: { fontSize: 22 },
        },
        styles: [{ styleId: "Normal", type: "paragraph", default: true }],
      },
      sourceTheme: theme,
      destinationTheme: theme,
      referencedStyleIds: [],
      materializeDefaultParagraphStyle: true,
    });

    expect(result.status).toBe("imported");
    if (result.status !== "imported") return;
    const alias = result.defaultParagraphStyleId;
    expect(alias).toBeDefined();
    expect(result.styles?.styles.at(-1)).toEqual({
      styleId: alias,
      type: "paragraph",
      pPr: { spaceBefore: 0, spaceAfter: 0, lineSpacing: 240, lineSpacingRule: "auto" },
      rPr: { fontSize: 22 },
    });
  });

  test("a target character style remains defined after accepting a comparison insertion", async () => {
    const targetStyles: StyleDefinitions = {
      ...destinationStyles,
      styles: [
        ...destinationStyles.styles,
        { styleId: "TargetCharacter", type: "character", rPr: { underline: { style: "double" } } },
      ],
    };
    const compared = await compareDocx(
      await documentWith({ styles: destinationStyles, text: "Before" }),
      await documentWith({
        styles: targetStyles,
        text: "After",
        formatting: { styleId: "TargetCharacter" },
      }),
      { author: "compare", timestamp: "2026-09-13T00:00:00.000Z" },
    );
    if (compared.isErr()) throw compared.error;

    expect(await stylesXml(compared.value.buffer)).toContain('w:styleId="TargetCharacter"');
    const pendingZip = await JSZip.loadAsync(compared.value.buffer);
    const pendingDocument = pendingZip.file("word/document.xml");
    if (!pendingDocument) throw new Error("Expected document.xml");
    const pendingDocumentXml = await pendingDocument.async("text");
    expect(pendingDocumentXml).not.toContain("<w:u ");

    const reviewer = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
    reviewer.acceptAll();
    expect(await stylesXml(await reviewer.toBuffer())).toContain('w:styleId="TargetCharacter"');
  });

  test("serializes a tracked style-context bridge with the old authored run formatting", async () => {
    const styles: StyleDefinitions = {
      docDefaults: { rPr: { fontFamily: { ascii: "Base Font", hAnsi: "Base Font" } } },
      styles: [
        { styleId: "Normal", type: "paragraph", default: true },
        {
          styleId: "Target",
          type: "paragraph",
          rPr: {
            fontFamily: {
              ascii: "Target Font",
              hAnsi: "Target Font",
              eastAsia: "Target Font",
              cs: "Target Font",
            },
            language: { val: "en-US", eastAsia: "en-US", bidi: "ar-SA" },
          },
        },
      ],
    };
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await documentWith({ styles, text: "Styled", formatting: { bold: true, boldCs: true } }),
    );
    const block = reviewer.snapshot().blocks.at(0);
    if (!block) throw new Error("Expected a source block");
    const result = reviewer.applyOperations([
      {
        id: "style-context",
        type: "replaceBlock",
        blockId: block.id,
        text: "Replacement",
        styleId: "Target",
      },
    ]);
    expect(result.skipped).toEqual([]);
    const pending = await reviewer.toBuffer();
    expect(await documentXml(pending)).toContain("<w:rPrChange");

    const rejected = await FolioDocxReviewer.fromBuffer(pending);
    rejected.rejectAll();
    expect(await documentXml(await rejected.toBuffer())).not.toContain('<w:rFonts ');
    expect(await documentXml(await rejected.toBuffer())).not.toContain('<w:lang ');
  });
});
