/**
 * A block custom XML element's own tracked records: the range markers
 * standing directly around it read as its insertion and deletion, and save
 * back around it; any other marker stays the captured markup it was.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, BlockCustomXml } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const paragraph = (text: string, id: string): string =>
  `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const wrap = (content: string): string =>
  `<w:customXml w:element="clause" w:uri="urn:folio:test">${content}</w:customXml>`;
const insStart = (id: number) =>
  `<w:customXmlInsRangeStart w:id="${id}" w:author="Ines" w:date="2026-03-01T00:00:00Z"/>`;
const insEnd = (id: number) => `<w:customXmlInsRangeEnd w:id="${id}"/>`;
const delStart = (id: number) =>
  `<w:customXmlDelRangeStart w:id="${id}" w:author="Dan" w:date="2026-03-02T00:00:00Z"/>`;
const delEnd = (id: number) => `<w:customXmlDelRangeEnd w:id="${id}"/>`;

const makeDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const parsedBlocks = async (body: string): Promise<BlockContent[]> =>
  (await parseDocx(await makeDocx(body), { preloadFonts: false })).package
    .document.content;

const elementOf = (blocks: readonly BlockContent[]): BlockCustomXml => {
  const element = blocks.find(
    (block): block is BlockCustomXml => block.type === "blockCustomXml",
  );
  if (element === undefined) throw new TypeError("No custom XML element.");
  return element;
};

const savedBody = async (body: string, viaEditor: boolean): Promise<string> => {
  const parsed = await parseDocx(await makeDocx(body), { preloadFonts: false });
  const document = viaEditor
    ? fromProseDoc(toProseDoc(parsed), parsed, {
        stylesheetSource: { type: "package" },
      })
    : parsed;
  const saved = await repackDocx(document, { updateModifiedDate: false });
  const xml =
    (await (await ZipArchive.read(saved)).entry("word/document.xml")?.text()) ??
    "";
  const start = xml.indexOf("<w:body>") + "<w:body>".length;
  return xml.slice(start, xml.indexOf("<w:sectPr", start));
};

describe("a block custom XML element's own records", () => {
  test("reads the insertion and deletion markers around it as its records", async () => {
    const blocks = await parsedBlocks(
      insStart(5) +
        delStart(6) +
        wrap(paragraph("Clause.", "11111111")) +
        delEnd(6) +
        insEnd(5),
    );
    expect(blocks.map(({ type }) => type)).toEqual(["blockCustomXml"]);
    const element = elementOf(blocks);
    expect(element.insertion).toEqual({
      info: { id: 5, author: "Ines", date: "2026-03-01T00:00:00Z" },
    });
    expect(element.deletion).toEqual({
      info: { id: 6, author: "Dan", date: "2026-03-02T00:00:00Z" },
    });
  });

  test("keeps a marker pairing with nothing beside the element as captured", async () => {
    const blocks = await parsedBlocks(
      insStart(5) +
        paragraph("Before.", "22222222") +
        wrap(paragraph("Clause.", "11111111")) +
        insEnd(5),
    );
    expect(blocks.map(({ type }) => type)).toEqual([
      "preservedBlock",
      "paragraph",
      "blockCustomXml",
      "preservedBlock",
    ]);
    expect(elementOf(blocks).insertion).toBeUndefined();
  });

  test("keeps a marker carrying markup the record has no field for", async () => {
    const blocks = await parsedBlocks(
      `<w:customXmlInsRangeStart w:id="5" w:author="Ines" w:custom="x"/>` +
        wrap(paragraph("Clause.", "11111111")) +
        insEnd(5),
    );
    expect(blocks.map(({ type }) => type)).toEqual([
      "preservedBlock",
      "blockCustomXml",
      "preservedBlock",
    ]);
  });

  test("keeps foreign same-named markers as captured markup", async () => {
    const blocks = await parsedBlocks(
      `<x:customXmlInsRangeStart xmlns:x="urn:extension" w:id="5" w:author="Ines"/>` +
        wrap(paragraph("Clause.", "11111111")) +
        `<x:customXmlInsRangeEnd xmlns:x="urn:extension" w:id="5"/>`,
    );
    expect(blocks.map(({ type }) => type)).toEqual([
      "preservedBlock",
      "blockCustomXml",
      "preservedBlock",
    ]);
  });

  test("keeps markers carrying foreign same-named attributes", async () => {
    const blocks = await parsedBlocks(
      `<w:customXmlInsRangeStart xmlns:x="urn:extension" w:id="5" w:author="Ines" x:date="keep"/>` +
        wrap(paragraph("Clause.", "11111111")) +
        insEnd(5),
    );
    expect(blocks.map(({ type }) => type)).toEqual([
      "preservedBlock",
      "blockCustomXml",
      "preservedBlock",
    ]);
  });

  test.each([
    { route: "the model", viaEditor: false },
    { route: "the editor projection", viaEditor: true },
  ])(
    "saves the records back around the element through $route",
    async ({ viaEditor }) => {
      const body =
        paragraph("Lead.", "33333333") +
        insStart(5) +
        delStart(6) +
        wrap(paragraph("Clause.", "11111111")) +
        delEnd(6) +
        insEnd(5);
      const saved = await savedBody(body, viaEditor);
      const opening = saved.indexOf("<w:customXml ");
      expect(saved.slice(0, opening)).toContain(
        '<w:customXmlInsRangeStart w:id="5" w:author="Ines" w:date="2026-03-01T00:00:00Z"/>' +
          '<w:customXmlDelRangeStart w:id="6" w:author="Dan" w:date="2026-03-02T00:00:00Z"/>',
      );
      expect(saved.slice(saved.indexOf("</w:customXml>"))).toStartWith(
        '</w:customXml><w:customXmlDelRangeEnd w:id="6"/><w:customXmlInsRangeEnd w:id="5"/>',
      );
      const reopened = await parsedBlocks(saved);
      const element = elementOf(reopened);
      expect(element.insertion?.info.id).toBe(5);
      expect(element.deletion?.info.id).toBe(6);
    },
  );

  test("reads records on an element in a table cell", async () => {
    const blocks = await parsedBlocks(
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>` +
        insStart(7) +
        wrap(paragraph("Cell clause.", "44444444")) +
        insEnd(7) +
        paragraph("", "55555555") +
        `</w:tc></w:tr></w:tbl>`,
    );
    const table = blocks.at(0);
    if (table?.type !== "table") throw new TypeError("No table.");
    const cell = table.rows.at(0)?.cells.at(0);
    if (cell === undefined) throw new TypeError("No cell.");
    expect(elementOf(cell.content).insertion?.info.id).toBe(7);
  });
});
