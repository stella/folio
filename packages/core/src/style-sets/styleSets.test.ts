import { describe, expect, test } from "bun:test";

import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "../docx/rezip";
import {
  extractDocumentStyleSet,
  extractDocumentStyleSetFromDocx,
  inspectDocumentStylesFromDocx,
} from "./extract";
import { createStellaStyleDocumentPreset } from "./stellaStyle";

describe("document style sets", () => {
  test("a document preset controls page shape and initial paragraph style", () => {
    const document = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const firstParagraph = document.package.document.content.at(0);

    expect(firstParagraph?.type).toBe("paragraph");
    expect(firstParagraph?.formatting?.styleId).toBe("BodyText");
    expect(document.package.document.finalSectionProperties?.pageWidth).toBe(11_906);
    expect(document.package.document.finalSectionProperties?.pageHeight).toBe(16_838);
    expect(document.package.styles?.styles.map((style) => style.styleId)).toContain(
      "ClauseHeading1",
    );
  });

  test("selected styles include their style and numbering dependencies", () => {
    const document = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const extracted = extractDocumentStyleSet(document, {
      name: "Clause styles",
      styleIds: ["ClauseParagraph1"],
      initialParagraphStyleId: "ClauseParagraph1",
    });

    expect(extracted.styles.styles.map((style) => style.styleId)).toEqual([
      "Normal",
      "BodyText",
      "ClauseParagraph1",
    ]);
    expect(extracted.numbering?.nums).toEqual([{ numId: 1, abstractNumId: 1 }]);
    expect(extracted.numbering?.abstractNums.map((numbering) => numbering.abstractNumId)).toEqual([
      1,
    ]);
  });

  test("rejects an initial paragraph style outside the selected style closure", () => {
    const document = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });

    expect(() =>
      extractDocumentStyleSet(document, {
        name: "Clause styles",
        styleIds: ["ClauseParagraph1"],
        initialParagraphStyleId: "ClauseHeading1",
      }),
    ).toThrow('Initial paragraph style "ClauseHeading1" is not present in the extracted set');
  });

  test("Stella Style uses the reference C-level clause hierarchy", () => {
    const styleSet = createStellaStyleDocumentPreset().styleSet;
    const clauseStyles = styleSet.styles.styles.filter((style) =>
      style.styleId.startsWith("Clause"),
    );

    expect(
      clauseStyles.map((style) => ({
        styleId: style.styleId,
        name: style.name,
        level: style.pPr?.numPr?.ilvl,
        bold: style.rPr?.bold ?? false,
        next: style.next,
      })),
    ).toEqual([
      {
        styleId: "ClauseHeading1",
        name: "Clause Heading",
        level: 0,
        bold: true,
        next: "ClauseParagraph1",
      },
      {
        styleId: "ClauseParagraph1",
        name: "Clause 1.1",
        level: 1,
        bold: false,
        next: "ClauseParagraph1",
      },
      {
        styleId: "ClauseParagraph2",
        name: "Clause (a)",
        level: 2,
        bold: false,
        next: "ClauseParagraph2",
      },
      {
        styleId: "ClauseParagraph3",
        name: "Clause (i)",
        level: 3,
        bold: false,
        next: "ClauseParagraph3",
      },
      {
        styleId: "ClauseParagraph4",
        name: "Clause (A)",
        level: 4,
        bold: false,
        next: "ClauseParagraph4",
      },
    ]);

    expect(styleSet.numbering?.abstractNums.at(0)?.levels).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ ilvl: 0, lvlText: "%1" }),
        expect.objectContaining({ ilvl: 1, lvlText: "%1.%2" }),
        expect.objectContaining({ ilvl: 2, lvlText: "(%3)" }),
        expect.objectContaining({ ilvl: 3, lvlText: "(%4)" }),
      ]),
    );

    const [mainClauseLevel, ordinaryClauseLevel] =
      styleSet.numbering?.abstractNums.at(0)?.levels ?? [];
    expect({
      mainMarkerLeft: mainClauseLevel?.pPr?.indentLeft,
      ordinaryMarkerLeft: ordinaryClauseLevel?.pPr?.indentLeft,
      mainHangingIndent: mainClauseLevel?.pPr?.indentFirstLine,
      ordinaryHangingIndent: ordinaryClauseLevel?.pPr?.indentFirstLine,
    }).toEqual({
      mainMarkerLeft: 567,
      ordinaryMarkerLeft: 567,
      mainHangingIndent: 567,
      ordinaryHangingIndent: 567,
    });
  });

  test("extraction strips document data and embedded-font relationships", () => {
    const document = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    document.package.properties = { title: "Confidential source title", creator: "Source user" };
    document.package.relationships = new Map([
      [
        "rId99",
        {
          id: "rId99",
          type: "attachedTemplate",
          target: "file:///private/template.dotx",
          targetMode: "External",
        },
      ],
    ]);
    const firstFont = document.package.fontTable?.fonts.at(0);
    if (!firstFont) {
      throw new Error("Expected legal preset to declare a font");
    }
    firstFont.embedRegular = "rIdFont";

    const extracted = extractDocumentStyleSet(document, { name: "Sanitized" });
    const serialized = JSON.stringify(extracted);

    expect(serialized).not.toContain("Confidential source title");
    expect(serialized).not.toContain("Source user");
    expect(serialized).not.toContain("private/template.dotx");
    expect(serialized).not.toContain("rIdFont");
  });

  test("a generated DOCX can be used as a style-set source", async () => {
    const source = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const buffer = await createDocx(source);
    const extracted = await extractDocumentStyleSetFromDocx(buffer, {
      name: "Imported legal",
      styleIds: ["FootnoteText"],
      initialParagraphStyleId: "Normal",
    });

    expect(extracted.styles.styles.map((style) => style.styleId)).toEqual([
      "Normal",
      "DefaultParagraphFont",
      "FootnoteText",
      "FootnoteTextChar",
    ]);
    expect(extracted.fontTable?.fonts.map((font) => font.name)).toEqual(["Arial", "Georgia"]);
    expect(extracted.numbering).toBeUndefined();
  });

  test("inspection exposes selectable style metadata without document content", async () => {
    const source = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const catalog = await inspectDocumentStylesFromDocx(await createDocx(source));

    expect(catalog.defaultParagraphStyleId).toBe("Normal");
    expect(catalog.styles.find((style) => style.styleId === "ClauseHeading1")).toEqual({
      styleId: "ClauseHeading1",
      name: "Clause Heading",
      type: "paragraph",
      role: "quick",
      dependencies: ["BodyText", "ClauseParagraph1"],
      numberingId: 1,
    });
  });
});
