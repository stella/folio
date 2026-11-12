/**
 * A complex field whose `begin` and `end` both sit inside a `w:hyperlink` is
 * one child of the link.
 *
 * A table of contents entry is one link over the heading text,
 * a tab and a `PAGEREF` page number. Read run by run, the page number's field
 * characters were loose runs in the link: the editor showed only the cached
 * number, and a save from the editor wrote the number without its field.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Hyperlink } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const run = (inner: string) => `<w:r>${inner}</w:r>`;
const begin = run('<w:fldChar w:fldCharType="begin"/>');
const separate = run('<w:fldChar w:fldCharType="separate"/>');
const end = run('<w:fldChar w:fldCharType="end"/>');
const instruction = (text: string) =>
  run(`<w:instrText xml:space="preserve">${text}</w:instrText>`);
const text = (value: string) => run(`<w:t>${value}</w:t>`);

const PAGE_REF = `${begin}${instruction(" PAGEREF _Toc1 \\h ")}${separate}${text("2")}${end}`;
const ENTRY = `<w:p><w:hyperlink w:anchor="_Toc1" w:history="1">${text("Summary")}${run("<w:tab/>")}${PAGE_REF}</w:hyperlink></w:p>`;

const parse = async (body: string): Promise<Document> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return parseDocx((await zip.generate({ compression: "store" })).buffer, {
    preloadFonts: false,
  });
};

const bodyOf = async (document: Document): Promise<string> => {
  const saved = await ZipArchive.read(
    await repackDocx(document, { updateModifiedDate: false }),
  );
  const xml = (await saved.entry("word/document.xml")?.text()) ?? "";
  return xml.slice(
    xml.indexOf("<w:body>") + "<w:body>".length,
    xml.indexOf("<w:sectPr"),
  );
};

const linkOf = (document: Document): Hyperlink => {
  const paragraph = document.package.document.content[0];
  const link =
    paragraph?.type === "paragraph"
      ? paragraph.content.find((item) => item.type === "hyperlink")
      : undefined;
  if (link?.type !== "hyperlink") throw new Error("No hyperlink");
  return link;
};

describe("a complex field inside a hyperlink", () => {
  test("is one child of the link", async () => {
    const link = linkOf(await parse(ENTRY));
    expect(link.children.map(({ type }) => type)).toEqual([
      "run",
      "run",
      "complexField",
    ]);
    const field = link.children[2];
    expect(field?.type === "complexField" && field.fieldType).toBe("PAGEREF");
  });

  test("saves inside the link as it was read", async () => {
    expect(await bodyOf(await parse(ENTRY))).toBe(ENTRY);
  });

  test("survives the editor's round trip inside the link", async () => {
    const document = await parse(ENTRY);
    const edited = fromProseDoc(
      toProseDoc(document, { styles: document.package.styles }),
      document,
      { stylesheetSource: { type: "package" } },
    );
    expect(linkOf(edited).children.map(({ type }) => type)).toEqual([
      "run",
      "run",
      "complexField",
    ]);
    expect(await bodyOf(edited)).toBe(ENTRY);
  });

  test("a field the link does not close stays as the runs it was read as", async () => {
    const body = `<w:p><w:hyperlink w:anchor="_Toc1" w:history="1">${text("Summary")}${begin}${instruction(" PAGEREF _Toc1 \\h ")}${separate}${text("2")}</w:hyperlink>${end}</w:p>`;
    const link = linkOf(await parse(body));
    expect(link.children.every(({ type }) => type === "run")).toBe(true);
    expect(await bodyOf(await parse(body))).toBe(body);
  });
});
