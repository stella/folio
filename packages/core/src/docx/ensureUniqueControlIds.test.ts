import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { ensureUniqueControlIds } from "./ensureUniqueControlIds";
import { createEmptyDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const control = (id: number, text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${String(id)}"/></w:sdtPr><w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt>`;

const packageOf = async (body: string, header?: string) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:mc="${MC}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  if (header !== undefined)
    zip.set(
      "word/header1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}">${header}</w:hdr>`,
    );
  return zip.generate({ compression: "store" });
};

const controlIds = async (docx: Uint8Array, path = "word/document.xml") =>
  [
    ...(
      (await (await ZipArchive.read(docx)).entry(path)?.text()) ?? ""
    ).matchAll(/<w:id w:val="(-?\d+)"\/>/gu),
  ].map(([, id]) => Number(id));

describe("unique content control ids", () => {
  test("the first control keeps its id and a later one gets an unused id", async () => {
    const docx = await packageOf(
      `<w:p>${control(5, "a")}</w:p><w:p>${control(5, "b")}</w:p><w:p>${control(1, "c")}</w:p>`,
    );
    const { docx: normalized, reminted } = await ensureUniqueControlIds(docx);
    expect(reminted).toBe(1);
    expect(await controlIds(normalized)).toEqual([5, 2, 1]);
    const xml =
      (await (
        await ZipArchive.read(normalized)
      )
        .entry("word/document.xml")
        ?.text()) ?? "";
    // Every control and its content stays.
    for (const text of ["a", "b", "c"])
      expect(xml).toContain(`<w:t>${text}</w:t>`);
  });

  test("a control repeated in another story part gets its own id", async () => {
    const docx = await packageOf(
      `<w:p>${control(5, "a")}</w:p>`,
      `<w:p>${control(5, "h")}</w:p>`,
    );
    const { docx: normalized } = await ensureUniqueControlIds(docx);
    expect(await controlIds(normalized)).toEqual([5]);
    expect(await controlIds(normalized, "word/header1.xml")).toEqual([1]);
  });

  test("alternate content restating a control keeps its id", async () => {
    const docx = await packageOf(
      `<w:p><mc:AlternateContent><mc:Choice Requires="w14">${control(5, "a")}</mc:Choice><mc:Fallback>${control(5, "a")}</mc:Fallback></mc:AlternateContent></w:p>`,
    );
    const result = await ensureUniqueControlIds(docx);
    expect(result.reminted).toBe(0);
    expect(result.docx).toBe(docx);
  });
});
