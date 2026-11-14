/**
 * The main document is whatever part the package's `officeDocument` relationship
 * names. A package that names another path still opens with its text, and an
 * edit lands in the part every reader of the saved package will read.
 */

import { describe, expect, mock, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createEmptyDocument } from "../utils/createDocument";
import {
  DOCUMENT,
  expectOneMainDocument,
  MAIN_TYPE,
  openedAsWritten,
  packageWithMainAt,
} from "./__tests__/mainDocumentAtAnotherPath";
import * as mainDocumentPart from "./mainDocumentPart";
import { mainDocumentPartPath } from "./mainDocumentPart";
import {
  addRelationship,
  createDocx,
  repackDocx,
  updateDocumentXml,
} from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";

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

describe("a save handed the package as its producer wrote it", () => {
  test("finds the part the package names", async () => {
    const zip = await ZipArchive.read(
      await packageWithMainAt("docs/main.xml", "/docs/main.xml"),
    );
    expect(await mainDocumentPartPath(zip)).toBe("docs/main.xml");
  });

  test("repackDocx writes one main document", async () => {
    const { document } = await openedAsWritten();
    await expectOneMainDocument(await repackDocx(document), "Authored text");
  });

  test("attemptSelectiveSave writes one main document", async () => {
    const { raw, document } = await openedAsWritten();
    await expectOneMainDocument(
      await attemptSelectiveSave(document, raw, {
        changedParaIds: new Set(["10000001"]),
        structuralChange: false,
        hasUntrackedChanges: false,
      }),
      "Authored text",
    );
  });

  test("updateDocumentXml replaces the main document", async () => {
    const { raw } = await openedAsWritten();
    await expectOneMainDocument(
      await updateDocumentXml(
        raw,
        DOCUMENT.replace("Authored text", "Replaced text"),
      ),
      "Replaced text",
    );
  });

  test("addRelationship adds to the main document's relationships", async () => {
    const { raw } = await openedAsWritten();
    const { buffer, rId } = await addRelationship(raw, {
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target: "https://example.com/",
      targetMode: "External",
    });
    await expectOneMainDocument(buffer, "Authored text");
    const relationships =
      (await (
        await ZipArchive.read(buffer)
      )
        .entry("word/_rels/document.xml.rels")
        ?.text()) ?? "";
    expect(relationships).toContain(`Id="${rId}"`);
  });

  test("a save that reads the package itself keeps the producer's main part", async () => {
    const { raw } = await openedAsWritten();
    const original = { ...mainDocumentPart };
    await mock.module("./mainDocumentPart", () => ({
      ...original,
      readPackageForSave: (bytes: ArrayBuffer | Uint8Array) =>
        ZipArchive.read(bytes),
    }));
    try {
      const saved = await ZipArchive.read(
        await updateDocumentXml(
          raw,
          DOCUMENT.replace("Authored text", "Replaced text"),
        ),
      );
      expect(await mainDocumentPartPath(saved)).toBe("word/document2.xml");
      expect(saved.has("word/document.xml")).toBe(true);
    } finally {
      await mock.module("./mainDocumentPart", () => original);
    }
  });
});
