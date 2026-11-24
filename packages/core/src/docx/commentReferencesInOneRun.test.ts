/**
 * One run may carry several `<w:commentReference>` children, each anchoring
 * its own comment. Every one is read, so an edit and a save keep each
 * comment's reference.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

const comment = (id: number, paraId: string, text: string): string =>
  `<w:comment w:id="${id}" w:author="Reviewer" w:initials="R" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="${paraId}"><w:r><w:annotationRef/></w:r><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;

const buildDocx = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      "</Types>",
  );
  zip.set(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdC" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/>` +
      "</Relationships>",
  );
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      '<w:p w14:paraId="1AAA0001"><w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="2"/>' +
      "<w:r><w:t>Shared words</w:t></w:r>" +
      '<w:commentRangeEnd w:id="1"/><w:commentRangeEnd w:id="2"/>' +
      '<w:r><w:commentReference w:id="1"/><w:commentReference w:id="2"/></w:r></w:p>' +
      '<w:p w14:paraId="1AAA0002"><w:r><w:t>Other paragraph.</w:t></w:r></w:p>' +
      "<w:sectPr/></w:body></w:document>",
  );
  zip.set(
    "word/comments.xml",
    `${XML_DECLARATION}<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
      comment(1, "2AAA0001", "First") +
      comment(2, "2AAA0002", "Second") +
      "</w:comments>",
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer as ArrayBuffer;
};

const references = async (buffer: ArrayBuffer): Promise<number[]> => {
  const xml =
    (await (
      await ZipArchive.read(buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return [...xml.matchAll(/<w:commentReference w:id="(\d+)"/gu)]
    .map((match) => Number(match[1]))
    .toSorted((left, right) => left - right);
};

describe("several comment references in one run", () => {
  test("the parser reads every reference in the run", async () => {
    const document = await parseDocx(await buildDocx(), {
      preloadFonts: false,
    });
    const first = document.package.document.content.at(0);
    if (first?.type !== "paragraph") throw new Error("no first paragraph");
    expect(
      first.content.flatMap((item) =>
        item.type === "commentReference" ? [item.id] : [],
      ),
    ).toEqual([1, 2]);
  });

  for (const session of ["canonical"] as const) {
    test(`${session}: an edit of the commented paragraph keeps both references`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocx(), {
        author: "Editor",
        session,
      });
      expect(reviewer.getComments().map(({ id }) => id)).toEqual([1, 2]);
      const commented = reviewer
        .snapshot()
        .blocks.find((block) => block.text === "Shared words");
      if (commented === undefined) throw new Error("no commented paragraph");
      const result = reviewer.applyOperations(
        [
          {
            id: "op1",
            type: "replaceInBlock",
            blockId: commented.id,
            find: "words",
            replace: "text",
          },
        ],
        { mode: "direct", undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      const saved = await reviewer.toBuffer();
      expect(await references(saved)).toEqual([1, 2]);
      const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
      expect(reopened.getComments().map(({ id }) => id)).toEqual([1, 2]);
    });
  }
});
