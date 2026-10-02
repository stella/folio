import { expect, test } from "bun:test";
import JSZip from "jszip";

import { panic } from "better-result";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";
import { parseSettings } from "./settingsParser";
import { updateEvenAndOddHeaders } from "./settingsHeaderFooterUpdate";
import { SETTINGS_CHILDREN } from "@stll/docx-core/schema";

for (const namespace of [
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]) {
  for (const prefix of ["s:", ""]) {
    test(`header/footer settings preserve foreign settings with ${namespace} and ${prefix || "default namespace"}`, () => {
      const binding = prefix ? `xmlns:s="${namespace}"` : `xmlns="${namespace}"`;
      const foreign = '<foreign:evenAndOddHeaders xmlns:foreign="urn:foreign"/>';
      const nested =
        '<foreign:extension xmlns:foreign="urn:foreign"><s:evenAndOddHeaders xmlns:s="urn:foreign"/></foreign:extension>';
      const xml = `<${prefix}settings ${binding}>${foreign}${nested}<!-- > --><?keep value?></${prefix}settings>`;
      for (const enabled of [true, false, undefined]) {
        const patched = updateEvenAndOddHeaders(xml, enabled) ?? panic("Expected settings patch");
        expect(parseSettings(patched).evenAndOddHeaders).toBe(enabled);
        expect(patched).toContain(foreign);
        expect(patched).toContain(nested);
        expect(patched).toContain("<!-- > --><?keep value?>");
        const toggled =
          updateEvenAndOddHeaders(patched, !enabled) ?? panic("Expected toggled patch");
        expect(parseSettings(toggled).evenAndOddHeaders).toBe(!enabled);
        expect(toggled).toContain(foreign);
        expect(toggled).toContain(nested);
      }
    });
  }
}

test("even header/footer setting materializes a missing part and survives subsequent saves", async () => {
  const initial = await createDocx(createEmptyDocument({ initialText: "Body" }));
  const zip = await JSZip.loadAsync(initial);
  zip.remove("word/settings.xml");
  const document = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
    preloadFonts: false,
  });
  document.package.settings = {
    ...document.package.settings,
    defaultTabStop: 720,
    evenAndOddHeaders: true,
  };
  const saved = await createDocx(document);
  const savedZip = await JSZip.loadAsync(saved);
  expect(await savedZip.file("[Content_Types].xml")?.async("text")).toContain(
    'PartName="/word/settings.xml"',
  );
  expect(await savedZip.file("word/_rels/document.xml.rels")?.async("text")).toContain(
    '/settings"',
  );
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(reopened.package.settings?.evenAndOddHeaders).toBe(true);
  if (!reopened.package.settings) panic("Expected settings");
  reopened.package.settings.evenAndOddHeaders = false;
  const disabled = await parseDocx(await createDocx(reopened), { preloadFonts: false });
  expect(disabled.package.settings?.evenAndOddHeaders).toBe(false);
});

test("self-closing settings roots expand through a validated XML splice", () => {
  const source =
    '<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" producer="kept"/>';
  const patched =
    updateEvenAndOddHeaders(source, true) ?? panic("Expected self-closing root patch");
  expect(parseSettings(patched).evenAndOddHeaders).toBe(true);
  expect(patched).toContain('producer="kept"');
  const removed = updateEvenAndOddHeaders(patched, undefined) ?? panic("Expected removed flag");
  expect(parseSettings(removed).evenAndOddHeaders).toBeUndefined();
});

test("evenAndOddHeaders is inserted at its generated CT_Settings sequence position", () => {
  const targetIndex = SETTINGS_CHILDREN.indexOf("evenAndOddHeaders");
  const followingChild = SETTINGS_CHILDREN.at(targetIndex + 1);
  if (targetIndex < 0 || followingChild === undefined) {
    panic("Expected evenAndOddHeaders to have a following CT_Settings child");
  }

  for (const namespace of [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
  ]) {
    const xml =
      `<w:settings xmlns:w="${namespace}">` +
      `<w:defaultTableStyle/><w:${followingChild}/><w:compat/></w:settings>`;
    const enabled = updateEvenAndOddHeaders(xml, true) ?? panic("Expected settings patch");
    const flagAt = enabled.indexOf("evenAndOddHeaders");
    const followingAt = enabled.indexOf(`<w:${followingChild}`);
    expect(parseSettings(enabled).evenAndOddHeaders).toBe(true);
    expect(flagAt).toBeGreaterThan(enabled.indexOf("<w:defaultTableStyle"));
    expect(flagAt).toBeLessThan(followingAt);

    const disabled =
      updateEvenAndOddHeaders(enabled, false) ?? panic("Expected existing flag update");
    expect(parseSettings(disabled).evenAndOddHeaders).toBe(false);
    expect(disabled.match(/evenAndOddHeaders/gu)).toHaveLength(1);
    expect(disabled.indexOf("evenAndOddHeaders")).toBeLessThan(
      disabled.indexOf(`<w:${followingChild}`),
    );
  }
});
