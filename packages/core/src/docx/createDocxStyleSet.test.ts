import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createStellaStyleDocumentPreset } from "../style-sets/stellaStyle";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

describe("createDocx definition parts", () => {
  test("exports every in-memory style for a generic empty document", async () => {
    const document = createEmptyDocument();
    const zip = await JSZip.loadAsync(await createDocx(document));
    const stylesXml = await zip.file("word/styles.xml")!.async("string");

    expect(
      [...stylesXml.matchAll(/w:styleId="(?<id>[^"]+)"/gu)].map((match) => match.groups!.id),
    ).toEqual([
      "Normal",
      "Title",
      "Subtitle",
      "Heading1",
      "Heading2",
      "Heading3",
      "Heading4",
      "Heading5",
      "Heading6",
      "TableNormal",
      "TableGrid",
      "Quote",
    ]);
  });

  test("writes each built-in under the w:name Word writes for it", async () => {
    // The style id is folio's to choose; the name is not. A consumer (Word's
    // own latent-style table, or `docx/builtInStyles.ts`) recognises a built-in
    // by this string, so these spellings are the file format's, taken from what
    // Microsoft Word writes in the public corpus. Word is not uniformly cased:
    // `heading 1`, `toc 1`, `footnote text` and `footer` are lowercase while
    // `Title`, `Body Text` and `Footnote Text Char` are not.
    const nameOf = (stylesXml: string, styleId: string): string | undefined =>
      new RegExp(`w:styleId="${styleId}"[^>]*>\\s*<w:name w:val="(?<name>[^"]*)"`, "u").exec(
        stylesXml,
      )?.groups?.["name"];

    const generic = await JSZip.loadAsync(await createDocx(createEmptyDocument()));
    const genericXml = await generic.file("word/styles.xml")!.async("string");
    expect({
      Normal: nameOf(genericXml, "Normal"),
      Title: nameOf(genericXml, "Title"),
      Subtitle: nameOf(genericXml, "Subtitle"),
      Heading1: nameOf(genericXml, "Heading1"),
      Heading4: nameOf(genericXml, "Heading4"),
      Quote: nameOf(genericXml, "Quote"),
    }).toEqual({
      Normal: "Normal",
      Title: "Title",
      Subtitle: "Subtitle",
      Heading1: "heading 1",
      Heading4: "heading 4",
      Quote: "Quote",
    });

    const stella = await JSZip.loadAsync(
      await createDocx(createEmptyDocument({ preset: createStellaStyleDocumentPreset() })),
    );
    const stellaXml = await stella.file("word/styles.xml")!.async("string");
    expect({
      Normal: nameOf(stellaXml, "Normal"),
      BodyText: nameOf(stellaXml, "BodyText"),
      Heading1: nameOf(stellaXml, "Heading1"),
      Heading6: nameOf(stellaXml, "Heading6"),
      TOC1: nameOf(stellaXml, "TOC1"),
      TOCHeading: nameOf(stellaXml, "TOCHeading"),
      ListParagraph: nameOf(stellaXml, "ListParagraph"),
      FootnoteText: nameOf(stellaXml, "FootnoteText"),
      FootnoteTextChar: nameOf(stellaXml, "FootnoteTextChar"),
      FootnoteReference: nameOf(stellaXml, "FootnoteReference"),
      EndnoteText: nameOf(stellaXml, "EndnoteText"),
      EndnoteReference: nameOf(stellaXml, "EndnoteReference"),
      Footer: nameOf(stellaXml, "Footer"),
      Hyperlink: nameOf(stellaXml, "Hyperlink"),
      TableNormal: nameOf(stellaXml, "TableNormal"),
      TableGrid: nameOf(stellaXml, "TableGrid"),
      DefaultParagraphFont: nameOf(stellaXml, "DefaultParagraphFont"),
    }).toEqual({
      Normal: "Normal",
      BodyText: "Body Text",
      Heading1: "heading 1",
      Heading6: "heading 6",
      TOC1: "toc 1",
      TOCHeading: "TOC Heading",
      ListParagraph: "List Paragraph",
      FootnoteText: "footnote text",
      FootnoteTextChar: "Footnote Text Char",
      FootnoteReference: "footnote reference",
      EndnoteText: "endnote text",
      EndnoteReference: "endnote reference",
      Footer: "footer",
      Hyperlink: "Hyperlink",
      TableNormal: "Normal Table",
      TableGrid: "Table Grid",
      DefaultParagraphFont: "Default Paragraph Font",
    });
  });

  test("materializes stella styles and all of their supported dependencies", async () => {
    const document = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const buffer = await createDocx(document);
    const zip = await JSZip.loadAsync(buffer);

    expect(zip.file("word/styles.xml")).not.toBeNull();
    expect(zip.file("word/numbering.xml")).not.toBeNull();
    expect(zip.file("word/fontTable.xml")).not.toBeNull();
    expect(zip.file("word/settings.xml")).not.toBeNull();

    const [contentTypes, relationships] = await Promise.all([
      zip.file("[Content_Types].xml")!.async("string"),
      zip.file("word/_rels/document.xml.rels")!.async("string"),
    ]);
    expect(contentTypes).toContain("/word/numbering.xml");
    expect(contentTypes).toContain("/word/fontTable.xml");
    expect(contentTypes).toContain("/word/settings.xml");
    expect(relationships).toContain('Target="numbering.xml"');
    expect(relationships).toContain('Target="fontTable.xml"');
    expect(relationships).toContain('Target="settings.xml"');

    const parsed = await parseDocx(buffer, { preloadFonts: false });
    expect(parsed.package.styles?.styles.map((style) => style.styleId)).toContain("ClauseHeading1");
    expect(parsed.package.numbering?.nums).toHaveLength(5);
    expect(parsed.package.fontTable?.fonts.map((font) => font.name)).toEqual(["Arial", "Georgia"]);
    expect(parsed.package.settings?.defaultTabStop).toBe(720);
    expect(parsed.package.document.content.at(0)?.formatting?.styleId).toBe("BodyText");
  });

  test("materializes a modeled theme for a new document", async () => {
    const preset = createStellaStyleDocumentPreset();
    preset.styleSet.theme = {
      name: "Imported theme",
      colorScheme: { accent1: "123456" },
      fontScheme: { minorFont: { latin: "Arial" } },
    };
    const buffer = await createDocx(createEmptyDocument({ preset }));
    const zip = await JSZip.loadAsync(buffer);

    expect(zip.file("word/theme/theme1.xml")).not.toBeNull();
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain(
      "/word/theme/theme1.xml",
    );
    expect(await zip.file("word/_rels/document.xml.rels")!.async("string")).toContain(
      'Target="theme/theme1.xml"',
    );

    const parsed = await parseDocx(buffer, { preloadFonts: false });
    expect(parsed.package.theme?.name).toBe("Imported theme");
    expect(parsed.package.theme?.colorScheme?.accent1).toBe("123456");
  });

  test("fails fast when a style references absent numbering", async () => {
    const preset = createStellaStyleDocumentPreset();
    preset.styleSet.numbering = undefined;

    await expect(createDocx(createEmptyDocument({ preset }))).rejects.toThrow(
      "Style references missing numbering definition 3",
    );
  });
});
