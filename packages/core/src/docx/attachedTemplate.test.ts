/**
 * `word/settings.xml` and its `.rels` part are preserved from the source
 * package. `w:attachedTemplate` points at a template outside the package
 * through a `TargetMode="External"` relationship; the document model has no
 * field for it, so a preserved copy would carry a reference folio can neither
 * read nor rewrite. Save filters both sides instead of copying them verbatim.
 *
 * The element is resolved by namespace URI plus local name, so the Transitional
 * and Strict profiles are both covered and a same-named foreign element stays.
 * Only the relationships the removed elements reference are dropped: the
 * settings part also carries mail-merge and transform relationships whose
 * `r:id` values must keep resolving.
 */
import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx, withoutAttachedTemplate } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";

const TRANSITIONAL_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const TRANSITIONAL_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships";

const SETTINGS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:settings xmlns:w="${TRANSITIONAL_W}" xmlns:r="${TRANSITIONAL_R}">` +
  '<w:defaultTabStop w:val="720"/><w:attachedTemplate r:id="rId1"/></w:settings>';

const relationship = (id: string, type: string, target: string): string =>
  `<Relationship Id="${id}"` +
  ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}"` +
  ` Target="${target}" TargetMode="External"/>`;

const SETTINGS_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  relationship("rId1", "attachedTemplate", "https://example.com/template.dotm") +
  relationship("rId2", "mailMergeSource", "https://example.com/recipients.csv") +
  "</Relationships>";

describe("attached template reference", () => {
  test("removes the Transitional element and only the relationship it names", () => {
    const filtered = withoutAttachedTemplate(SETTINGS_XML, SETTINGS_RELS);

    expect(filtered.settingsXml).not.toContain("attachedTemplate");
    expect(filtered.settingsXml).toContain('w:defaultTabStop w:val="720"');
    // A mail-merge source is a separate settings relationship; its `r:id` must
    // keep resolving after the template reference goes.
    expect(filtered.relsXml).not.toContain("template.dotm");
    expect(filtered.relsXml).toContain('Id="rId2"');
    expect(filtered.relsXml).toContain("recipients.csv");
  });

  test("removes the Strict element", () => {
    const strict =
      `<settings xmlns="${STRICT_W}" xmlns:r="${STRICT_R}">` +
      '<defaultTabStop w:val="720" xmlns:w="' +
      STRICT_W +
      '"/><attachedTemplate r:id="rId1"/></settings>';

    expect(withoutAttachedTemplate(strict, undefined).settingsXml).not.toContain(
      "attachedTemplate",
    );
  });

  test("keeps a same-named element from a foreign namespace", () => {
    const foreign =
      `<w:settings xmlns:w="${TRANSITIONAL_W}" xmlns:r="${TRANSITIONAL_R}" xmlns:x="urn:example:ext">` +
      '<x:attachedTemplate r:id="rId1"/></w:settings>';

    // No WordprocessingML element matched, so neither part is rewritten.
    expect(withoutAttachedTemplate(foreign, SETTINGS_RELS)).toEqual({
      settingsXml: undefined,
      relsXml: undefined,
    });
  });

  test("removes a paired element and leaves an unrelated part untouched", () => {
    const paired =
      `<w:settings xmlns:w="${TRANSITIONAL_W}" xmlns:r="${TRANSITIONAL_R}">` +
      '<w:attachedTemplate r:id="rId1"></w:attachedTemplate></w:settings>';

    expect(withoutAttachedTemplate(paired, undefined).settingsXml).toBe(
      `<w:settings xmlns:w="${TRANSITIONAL_W}" xmlns:r="${TRANSITIONAL_R}"></w:settings>`,
    );
    expect(
      withoutAttachedTemplate('<w:settings xmlns:w="' + TRANSITIONAL_W + '"/>', undefined),
    ).toEqual({ settingsXml: undefined, relsXml: undefined });
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
    expect(rels).toContain("recipients.csv");
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
