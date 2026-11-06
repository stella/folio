/**
 * An edit to a text box written with a fallback rewrites both branches. The
 * paragraphs carry `w14:paraId`, which the captured branches do not declare.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="w14"';

// No paragraph id of its own, so no `w14` in the captured branches; the save gives it one.
const box = `<w:p><w:r><w:t>Boxed words</w:t></w:r></w:p>`;

const ALTERNATE = `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${box}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect><v:textbox><w:txbxContent>${box}</w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent>`;

const open = async () => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="10000001"><w:r>${ALTERNATE}</w:r><w:r><w:t>Host</w:t></w:r></w:p><w:p w14:paraId="10000003"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`,
  );
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer);
};

describe("a text box with a fallback branch", () => {
  for (const mode of ["direct", "tracked-changes"] as const) {
    test(`saves after an edit to its text (${mode})`, async () => {
      const reviewer = await open();
      const boxed = reviewer
        .getContent()
        .find(({ text }) => text === "Boxed words");
      if (!boxed) throw new Error("The text box's paragraph is missing");
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode,
          operations: [
            {
              id: "edit",
              type: "replaceInBlock",
              blockId: boxed.id,
              find: "words",
              replace: "text",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.applied).toHaveLength(1);
      const saved = await reviewer.toBuffer();
      const xml =
        (await (
          await ZipArchive.read(saved)
        )
          .entry("word/document.xml")
          ?.text()) ?? "";
      // Both branches carry the edit.
      expect(xml.match(/text<\/w:t>/gu)).toHaveLength(2);
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      reopened.acceptAll();
      expect(reopened.getContent().map(({ text }) => text)).toContain(
        "Boxed text",
      );
    });
  }
});
