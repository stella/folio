/**
 * A comment whose only anchor is its `w:commentReference`, as a file holds when
 * the text it covered was deleted, survives every edit that leaves the
 * reference alone, and goes with the reference.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const open = async () => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="10000001"><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t>Host</w:t></w:r></w:p><w:p w14:paraId="10000002"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`,
  );
  zip.set(
    "word/comments.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:comments ${NS}><w:comment w:id="1" w:author="A" w:initials="A"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  const types = (await zip.entry("[Content_Types].xml")?.text()) ?? "";
  zip.set(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
    ),
  );
  const relationships =
    (await zip.entry("word/_rels/document.xml.rels")?.text()) ?? "";
  zip.set(
    "word/_rels/document.xml.rels",
    relationships.replace(
      "</Relationships>",
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
    ),
  );
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer);
};

const edit = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  operation: Record<string, unknown>,
) =>
  reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [{ id: "edit", ...operation }] } as never,
    { undefinedReferences: "refuse" },
  );

const savedCommentTexts = async (reviewer: FolioDocxReviewer) =>
  (await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()))
    .getComments()
    .map(({ text }) => text);

describe.each(["direct", "tracked-changes"] as const)(
  "a comment anchored by its reference alone (%s)",
  (mode) => {
    test("survives an edit elsewhere", async () => {
      const reviewer = await open();
      expect(
        edit(reviewer, mode, {
          type: "replaceInBlock",
          blockId: "10000002",
          find: "Tail",
          replace: "End",
        }).applied,
      ).toHaveLength(1);
      expect(reviewer.getComments()).toHaveLength(1);
      expect(await savedCommentTexts(reviewer)).toEqual(["Note"]);
    });

    test("survives an edit in its own paragraph", async () => {
      const reviewer = await open();
      edit(reviewer, mode, {
        type: "replaceInBlock",
        blockId: "10000001",
        find: "Host",
        replace: "Guest",
      });
      expect(await savedCommentTexts(reviewer)).toEqual(["Note"]);
    });

    test("survives accepting and rejecting other changes", async () => {
      for (const resolve of ["acceptAll", "rejectAll"] as const) {
        const reviewer = await open();
        edit(reviewer, mode, {
          type: "insertAfterBlock",
          blockId: "10000002",
          text: "More",
        });
        reviewer[resolve]();
        expect(await savedCommentTexts(reviewer)).toEqual(["Note"]);
      }
    });
  },
);

test("a comment anchored by its reference goes with its paragraph", async () => {
  const reviewer = await open();
  edit(reviewer, "direct", { type: "deleteBlock", blockId: "10000001" });
  expect(await savedCommentTexts(reviewer)).toEqual([]);
});
