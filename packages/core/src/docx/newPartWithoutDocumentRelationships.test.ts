/**
 * A document that refers to no other part may come without
 * `word/_rels/document.xml.rels`. An edit that needs a new part (a list's
 * numbering, a comment) starts the relationships part on save, and the
 * reopened package has what the edit added.
 */
import { describe, expect, test } from "bun:test";

import { paragraphNumberingReferenceId } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioDocumentOperation } from "../ai-edits/types";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DOCUMENT_RELS = "word/_rels/document.xml.rels";

const packageWithoutDocumentRelationships = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="11111111"><w:r><w:t>Opening words.</w:t></w:r></w:p>` +
      `</w:body></w:document>`,
  );
  return (await zip.generate()).buffer;
};

const ADDED = "Added words.";

const EDITS: Record<string, FolioDocumentOperation> = {
  list: {
    id: "list",
    type: "insertAfterBlock",
    blockId: "11111111",
    text: ADDED,
    numbering: { kind: "newList", format: "numbered" },
  },
  comment: {
    id: "comment",
    type: "commentOnBlock",
    blockId: "11111111",
    comment: { text: ADDED },
  },
};

describe("a package without document relationships", () => {
  test("saves unedited without gaining them", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await packageWithoutDocumentRelationships(),
    );
    const saved = await ZipArchive.read(await reviewer.toBuffer());
    expect(saved.entry(DOCUMENT_RELS)).toBeUndefined();
  });

  for (const mode of ["direct", "tracked-changes"] as const) {
    for (const [name, operation] of Object.entries(EDITS)) {
      test(`saves a ${name} added in ${mode} mode`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await packageWithoutDocumentRelationships(),
          { author: "Agent" },
        );
        const result = reviewer.applyOperations([operation], {
          undefinedReferences: "refuse",
          mode,
        });
        expect(result.applied).toHaveLength(1);

        const saved = await reviewer.toBuffer();
        const zip = await ZipArchive.read(saved);
        const relationships = (await zip.entry(DOCUMENT_RELS)?.text()) ?? "";
        const reopened = await FolioDocxReviewer.fromBuffer(saved);
        if (name === "comment") {
          expect(relationships).toContain(RELATIONSHIP_TYPES.comments);
          expect(reopened.getComments().map(({ text }) => text)).toContain(
            ADDED,
          );
          return;
        }
        expect(reopened.getContent().map(({ text }) => text)).toContain(ADDED);
        expect(relationships).toContain(RELATIONSHIP_TYPES.numbering);
        const parsed = await parseDocx(saved, {
          preloadFonts: false,
          detectVariables: false,
        });
        const added = parsed.package.document.content.find(
          (block) =>
            block.type === "paragraph" &&
            JSON.stringify(block.content).includes(ADDED),
        );
        if (added?.type !== "paragraph") throw new Error("no added paragraph");
        const numId = paragraphNumberingReferenceId(added.formatting?.numPr);
        expect(numId).toBeDefined();
        expect(
          parsed.package.numbering?.nums.some((num) => num.numId === numId),
        ).toBe(true);
      });
    }
  }
});
