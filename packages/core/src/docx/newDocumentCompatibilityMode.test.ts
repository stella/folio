import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { resolveJustificationCompatibility } from "../layout-engine/justificationCompatibility";
import { resolveTableIndentCompatibility } from "../layout-engine/tableIndentCompatibility";
import { createStellaStyleDocumentPreset } from "../style-sets/stellaStyle";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx, createEmptyDocx, repackDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const COMPAT_URI = "http://schemas.microsoft.com/office/word";

const settingsXml = (body: string): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/>${body}</w:settings>`;

const packageWithSettings = async (settings: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/settings.xml", settings);
  const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    ),
  );
  const types = await zip.file("[Content_Types].xml")!.async("string");
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const savedSettings = async (buffer: ArrayBuffer): Promise<string | undefined> =>
  (await JSZip.loadAsync(buffer)).file("word/settings.xml")?.async("string");

describe("new document compatibility mode", () => {
  test("a new document declares compatibility mode 15 to layout", () => {
    for (const document of [
      createEmptyDocument(),
      createEmptyDocument({ preset: createStellaStyleDocumentPreset() }),
    ]) {
      const mode = document.package.settings?.compatibilityMode;
      expect(mode).toBe(15);
      expect(resolveJustificationCompatibility(mode)).toBeUndefined();
      expect(resolveTableIndentCompatibility(mode)).toBeUndefined();
    }
  });

  test("a style set that states a compatibility mode keeps it", () => {
    const preset = createStellaStyleDocumentPreset();
    preset.styleSet.settings = { defaultTabStop: 720, compatibilityMode: 14 };

    expect(createEmptyDocument({ preset }).package.settings?.compatibilityMode).toBe(14);
  });

  test("a new package writes the mode and reads it back", async () => {
    const buffer = await createDocx(createEmptyDocument());
    const settings = await savedSettings(buffer);

    expect(settings).toContain(
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="${COMPAT_URI}" w:val="15"/>` +
        `<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="${COMPAT_URI}" w:val="1"/>` +
        `<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="${COMPAT_URI}" w:val="1"/>` +
        "</w:compat>",
    );
    const parsed = await parseDocx(buffer, { preloadFonts: false });
    expect(parsed.package.settings?.compatibilityMode).toBe(15);
  });

  test("a new package in an older mode writes only the mode", async () => {
    const preset = createStellaStyleDocumentPreset();
    preset.styleSet.settings = { defaultTabStop: 720, compatibilityMode: 14 };
    const settings = await savedSettings(await createDocx(createEmptyDocument({ preset })));

    expect(settings).toContain(
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="${COMPAT_URI}" w:val="14"/></w:compat>`,
    );
  });

  test("a loaded package without a compatibility mode saves without one", async () => {
    const source = settingsXml("");
    const document = await parseDocx(await packageWithSettings(source), { preloadFonts: false });

    expect(document.package.settings?.compatibilityMode).toBeUndefined();
    expect(resolveJustificationCompatibility(document.package.settings?.compatibilityMode)).toEqual(
      { type: "legacy" },
    );
    expect(await savedSettings(await repackDocx(document, { updateModifiedDate: false }))).toBe(
      source,
    );
  });

  test("a loaded package keeps its authored compatibility settings", async () => {
    const source = settingsXml(
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="${COMPAT_URI}" w:val="14"/>` +
        `<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="${COMPAT_URI}" w:val="1"/>` +
        "</w:compat>",
    );
    const document = await parseDocx(await packageWithSettings(source), { preloadFonts: false });

    expect(document.package.settings?.compatibilityMode).toBe(14);
    expect(await savedSettings(await repackDocx(document, { updateModifiedDate: false }))).toBe(
      source,
    );
  });
});
