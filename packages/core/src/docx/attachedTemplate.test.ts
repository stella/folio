/**
 * `word/settings.xml` and its `.rels` part are preserved from the source
 * package. `w:attachedTemplate` points at a template outside the package
 * through a `TargetMode="External"` relationship; the document model has no
 * field for it, so a preserved copy would carry a reference folio can neither
 * read nor rewrite. Save filters both sides instead of copying them verbatim.
 */
import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import {
  createEmptyDocx,
  removeAttachedTemplateElement,
  removeExternalRelationships,
  repackDocx,
} from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";

const SETTINGS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<w:defaultTabStop w:val="720"/><w:attachedTemplate r:id="rId1"/></w:settings>';

const SETTINGS_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1"' +
  ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"' +
  ' Target="https://example.com/template.dotm" TargetMode="External"/>' +
  "</Relationships>";

describe("attached template reference", () => {
  test("removeAttachedTemplateElement drops the element in either form", () => {
    expect(removeAttachedTemplateElement(SETTINGS_XML)).not.toContain("attachedTemplate");
    expect(removeAttachedTemplateElement(SETTINGS_XML)).toContain('w:defaultTabStop w:val="720"');
    expect(
      removeAttachedTemplateElement(
        '<w:settings><w:attachedTemplate r:id="rId1"></w:attachedTemplate></w:settings>',
      ),
    ).toBe("<w:settings></w:settings>");
    expect(
      removeAttachedTemplateElement('<w:settings><x:attachedTemplate r:id="rId1"/></w:settings>'),
    ).toBe("<w:settings></w:settings>");
  });

  test("removeExternalRelationships keeps package-internal entries", () => {
    expect(removeExternalRelationships(SETTINGS_RELS)).not.toContain('Id="rId1"');

    const internal =
      '<Relationships><Relationship Id="rId2" Type="urn:t" Target="styles.xml"/></Relationships>';
    expect(removeExternalRelationships(internal)).toBe(internal);
  });

  const seedSource = async (): Promise<ArrayBuffer> => {
    const seeded = new JSZip();
    await seeded.loadAsync(await createEmptyDocx());
    seeded.file("word/settings.xml", SETTINGS_XML);
    seeded.file("word/_rels/settings.xml.rels", SETTINGS_RELS);
    return seeded.generateAsync({ type: "arraybuffer" });
  };

  const expectFiltered = async (out: ArrayBuffer): Promise<void> => {
    const zip = await JSZip.loadAsync(out);
    const settings = await zip.file("word/settings.xml")!.async("text");
    expect(settings).not.toContain("attachedTemplate");
    expect(settings).toContain('w:defaultTabStop w:val="720"');

    const rels = await zip.file("word/_rels/settings.xml.rels")!.async("text");
    expect(rels).not.toContain("template.dotm");
    expect(rels).not.toContain('TargetMode="External"');
  };

  test("a full repack drops the reference and its relationship", async () => {
    const doc = await parseDocx(await seedSource(), { preloadFonts: false });

    await expectFiltered(await repackDocx(doc, { updateModifiedDate: false }));
  });

  test("a selective save drops them the same way", async () => {
    // Both save paths must emit the same package for the same source.
    const source = await seedSource();
    const doc = await parseDocx(source, { preloadFonts: false });
    const out = await attemptSelectiveSave(doc, source, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(out).not.toBeNull();
    await expectFiltered(out!);
  });
});
