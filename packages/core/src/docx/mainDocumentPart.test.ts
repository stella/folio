/**
 * The main document is whatever part the package's `officeDocument` relationship
 * names. A package that names another path still opens with its text, and an
 * edit lands in the part every reader of the saved package will read.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "./rezip";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const DOCUMENT = `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="10000001"><w:r><w:t>Authored text</w:t></w:r></w:p></w:body></w:document>`;
const MAIN_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/** The standard package, with its main document moved to `path`. */
const packageWithMainAt = async (
  path: string,
  target = path,
): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  const relationships = await zip.entry("word/_rels/document.xml.rels")?.text();
  const packageRelationships = await zip.entry("_rels/.rels")?.text();
  const contentTypes = await zip.entry("[Content_Types].xml")?.text();
  if (!relationships || !packageRelationships || !contentTypes) {
    throw new Error("The base package is missing a part");
  }
  zip.delete("word/document.xml");
  zip.delete("word/_rels/document.xml.rels");
  zip.set(path, DOCUMENT);
  const slash = path.lastIndexOf("/");
  zip.set(
    `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`,
    relationships,
  );
  zip.set(
    "_rels/.rels",
    packageRelationships.replace(
      /Target="\/?word\/document\.xml"/u,
      `Target="${target}"`,
    ),
  );
  zip.set(
    "[Content_Types].xml",
    contentTypes.replace(
      /PartName="\/word\/document\.xml"/u,
      `PartName="/${path}"`,
    ),
  );
  return (await zip.generate()).slice().buffer;
};

describe("a main document at another path", () => {
  test.each([
    ["word/document2.xml", "word/document2.xml"],
    ["word/document22.xml", "/word/document22.xml"],
    ["word/blah_blah.xml", "word/blah_blah.xml"],
    ["docs/main.xml", "docs/main.xml"],
  ])("opens with its text from %s", async (path, target) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await packageWithMainAt(path, target),
    );
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "Authored text",
    ]);
  });

  for (const mode of ["direct", "tracked-changes"] as const) {
    test(`saves an edit where the package says the document is (${mode})`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await packageWithMainAt("word/document2.xml"),
      );
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode,
          operations: [
            {
              id: "edit",
              type: "insertAfterBlock",
              blockId: "10000001",
              text: "Added",
              numbering: { kind: "newList", format: "numbered" },
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.applied).toHaveLength(1);
      const saved = await ZipArchive.read(await reviewer.toBuffer());
      const packageRelationships =
        (await saved.entry("_rels/.rels")?.text()) ?? "";
      const named =
        /officeDocument"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*officeDocument"/u.exec(
          packageRelationships,
        );
      const main = (named?.[1] ?? named?.[2] ?? "").replace(/^\//u, "");
      expect(main).toBe("word/document.xml");
      expect(saved.has("word/document2.xml")).toBe(false);
      expect(await saved.entry(main)?.text()).toContain("Added");
      const contentTypes =
        (await saved.entry("[Content_Types].xml")?.text()) ?? "";
      expect(contentTypes).toContain(
        `PartName="/word/document.xml" ContentType="${MAIN_TYPE}"`,
      );
      expect(contentTypes).not.toContain("document2.xml");
    });
  }

  test("keeps a part the relationship does not name out of the way", async () => {
    const zip = await ZipArchive.read(
      await packageWithMainAt("word/document2.xml"),
    );
    zip.set(
      "word/document.xml",
      DOCUMENT.replace("Authored text", "Stale leftover"),
    );
    const reviewer = await FolioDocxReviewer.fromBuffer(
      (await zip.generate()).slice().buffer,
    );
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "Authored text",
    ]);
  });

  test("leaves a package whose named part is no document alone", async () => {
    const bytes = await createDocx(createEmptyDocument());
    const zip = await ZipArchive.read(bytes);
    const packageRelationships = await zip.entry("_rels/.rels")?.text();
    if (!packageRelationships) throw new Error("No package relationships");
    zip.set("theme/theme/themeManager.xml", "<a/>");
    zip.set(
      "_rels/.rels",
      packageRelationships.replace(
        /Target="\/?word\/document\.xml"/u,
        'Target="theme/theme/themeManager.xml"',
      ),
    );
    const moved = await FolioDocxReviewer.fromBuffer(
      (await zip.generate()).slice().buffer,
    );
    expect(moved.getContent()).toHaveLength(1);
  });
});
