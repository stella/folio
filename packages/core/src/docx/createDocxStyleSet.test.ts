import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createStellaStyleDocumentPreset } from "../style-sets/stellaStyle";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

describe("createDocx definition parts", () => {
  test("writes a Word-compatible extended-properties application version", async () => {
    const zip = await JSZip.loadAsync(await createDocx(createEmptyDocument()));
    const appPropertiesXml = await zip.file("docProps/app.xml")!.async("string");

    expect(appPropertiesXml).toContain("<AppVersion>1.0000</AppVersion>");
    expect(appPropertiesXml).not.toContain("<AppVersion>1.0.0</AppVersion>");
  });

  test("exports every in-memory style for a generic empty document", async () => {
    const document = createEmptyDocument();
    const zip = await JSZip.loadAsync(await createDocx(document));
    const stylesXml = await zip.file("word/styles.xml")!.async("string");

    expect(
      [...stylesXml.matchAll(/w:styleId="(?<id>[^"]+)"/gu)].map((match) => match.groups!.id),
    ).toEqual(["Normal", "Title", "Subtitle", "Heading1", "Heading2", "Heading3", "Heading4"]);
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
