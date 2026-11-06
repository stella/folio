/**
 * The schema lets a table row hold no cells at all (`CT_Row` takes zero or more
 * cell content items). A document with one opens, edits and saves.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const cell = (text: string) =>
  `<w:tc><w:tcPr/><w:p w14:paraId="1000000${text}"><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

const ROWS = {
  "nothing at all": "<w:tr/>",
  "only properties": '<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr></w:tr>',
  "only a w:gridAfter":
    '<w:tr><w:trPr><w:gridAfter w:val="2"/></w:trPr></w:tr>',
  "only a w:gridBefore":
    '<w:tr><w:trPr><w:gridBefore w:val="2"/></w:trPr></w:tr>',
} as const;

const docx = async (row: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${cell("1")}${cell("2")}</w:tr>${row}<w:tr>${cell("3")}${cell("4")}</w:tr></w:tbl><w:p w14:paraId="10000009"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`,
  );
  return (await zip.generate()).slice().buffer;
};

const authoredText = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getContent()
    .map(({ text }) => text)
    .filter((text) => text !== "");

const SOURCE = ["1", "2", "3", "4", "Tail"];

const savedRowCount = async (buffer: ArrayBuffer) => {
  const xml =
    (await (
      await ZipArchive.read(buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return (xml.match(/<w:tr[ />]/gu) ?? []).length;
};

describe.each(Object.entries(ROWS))("a table row with %s", (_name, row) => {
  test("opens", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await docx(row));
    expect(authoredText(reviewer)).toEqual(SOURCE);
  });

  test("saves untouched, and keeps its place", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await docx(row));
    const saved = await reviewer.toBuffer();
    expect(await savedRowCount(saved)).toBe(3);
    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    expect(authoredText(reopened)).toEqual(SOURCE);
  });

  test.each(["direct", "tracked-changes"] as const)(
    "saves after edits around it (%s)",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await docx(row));
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode,
          operations: [
            {
              id: "edit",
              type: "replaceInBlock",
              blockId: "10000001",
              find: "1",
              replace: "9",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.applied).toHaveLength(1);
      const saved = await reviewer.toBuffer();
      expect(await savedRowCount(saved)).toBe(3);
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      reopened.acceptAll();
      expect(authoredText(reopened)).toEqual(["9", "2", "3", "4", "Tail"]);
    },
  );

  test("stays through a merge and a column edit elsewhere", async () => {
    const merged = await FolioDocxReviewer.fromBuffer(await docx(row));
    expect(
      merged.applyDocumentOperations(
        {
          version: 1,
          mode: "direct",
          operations: [
            {
              id: "merge",
              type: "mergeTableCells",
              blockId: "10000003",
              endBlockId: "10000004",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      ).applied,
    ).toHaveLength(1);
    expect(await savedRowCount(await merged.toBuffer())).toBe(3);

    const narrowed = await FolioDocxReviewer.fromBuffer(await docx(row));
    expect(
      narrowed.applyDocumentOperations(
        {
          version: 1,
          mode: "direct",
          operations: [
            { id: "column", type: "deleteTableColumn", blockId: "10000002" },
          ],
        },
        { undefinedReferences: "refuse" },
      ).applied,
    ).toHaveLength(1);
    expect(await savedRowCount(await narrowed.toBuffer())).toBe(3);
  });
});
