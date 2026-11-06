/**
 * A save repairs what it broke. A relationship, an `r:embed` or a content-type
 * override that named nothing in the source is the source's own, and an edit
 * elsewhere leaves it, and the part holding it, as they were.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { createEmptyDocument } from "../../utils/createDocument";
import { createDocx } from "../rezip";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const RELATIONSHIPS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const PICTURE = (id: string) =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="P"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${id}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const open = async () => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/header1.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${NS}><w:p>${PICTURE("rIdGone")}</w:p></w:hdr>`,
  );
  zip.set(
    "word/_rels/header1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdGone" Type="${RELATIONSHIPS}/image" Target="media/absent.png"/></Relationships>`,
  );
  const types = (await zip.entry("[Content_Types].xml")?.text()) ?? "";
  zip.set(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/absent.xml" ContentType="application/xml"/></Types>',
    ),
  );
  const relationships =
    (await zip.entry("word/_rels/document.xml.rels")?.text()) ?? "";
  zip.set(
    "word/_rels/document.xml.rels",
    relationships.replace(
      "</Relationships>",
      `<Relationship Id="rIdHeader" Type="${RELATIONSHIPS}/header" Target="header1.xml"/><Relationship Id="rIdAbsent" Type="${RELATIONSHIPS}/image" Target="media/absent2.png"/></Relationships>`,
    ),
  );
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="10000001"><w:r><w:t>Text</w:t></w:r></w:p><w:p w14:paraId="10000002"><w:r><w:t>Tail</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate()).slice().buffer;
};

const part = async (buffer: ArrayBuffer, path: string) =>
  (await (await ZipArchive.read(buffer)).entry(path)?.text()) ?? "";

describe.each(["direct", "tracked-changes"] as const)(
  "an edit elsewhere in a package with dangling references (%s)",
  (mode) => {
    const edited = async () => {
      const source = await open();
      const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0));
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode,
          operations: [
            // Numbers the paragraph, so the save is a full repack.
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
      return { source, saved: await reviewer.toBuffer() };
    };

    test("keeps the header and its relationships as they were", async () => {
      const { source, saved } = await edited();
      for (const path of ["word/header1.xml", "word/_rels/header1.xml.rels"]) {
        expect(await part(saved, path)).toBe(await part(source, path));
      }
    });

    test("keeps a dangling relationship of the document", async () => {
      const { saved } = await edited();
      expect(await part(saved, "word/_rels/document.xml.rels")).toContain(
        "media/absent2.png",
      );
    });

    test("keeps a content-type override for a part it never had", async () => {
      const { saved } = await edited();
      expect(await part(saved, "[Content_Types].xml")).toContain(
        "/word/absent.xml",
      );
    });
  },
);

test("a relationship to an entry that keeps its name's escapes is not dangling", async () => {
  const zip = await ZipArchive.read(await open());
  zip.set("word/media/my%20image.png", new Uint8Array([1, 2, 3]));
  const relationships =
    (await zip.entry("word/_rels/document.xml.rels")?.text()) ?? "";
  zip.set(
    "word/_rels/document.xml.rels",
    relationships.replace(
      "</Relationships>",
      `<Relationship Id="rIdEscaped" Type="${RELATIONSHIPS}/image" Target="media/my%20image.png"/></Relationships>`,
    ),
  );
  const reviewer = await FolioDocxReviewer.fromBuffer(
    (await zip.generate()).slice().buffer,
  );
  reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "direct",
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
  const saved = await reviewer.toBuffer();
  expect(await part(saved, "word/_rels/document.xml.rels")).toContain(
    "my%20image.png",
  );
  expect((await ZipArchive.read(saved)).has("word/media/my%20image.png")).toBe(
    true,
  );
});
